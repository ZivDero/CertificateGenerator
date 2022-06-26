using CertificateGenerator.ViewModel;
using CG.Web.MegaApiClient;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using iText.IO.Source;
using MimeKit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Message = Google.Apis.Gmail.v1.Data.Message;
using MessageBox = System.Windows.Forms.MessageBox;

namespace CertificateGenerator.Handlers
{
    public class SmtpHandler
    {
        private MessageProperties properties;

        private GmailService service;
        UserCredential credential;
        static string[] Scopes = { GmailService.Scope.GmailSend };
        static string ApplicationName = "Certificate Generator";

        private List<string> attachments = new List<string>();
        private MegaApiClient megaClient = new MegaApiClient();
        
        private readonly MainViewModel viewModel;

        public SmtpHandler(MainViewModel vm)
        {
            viewModel = vm;
        }

        public void AddAttachment(string path)
        {
            attachments.Add(path);
        }

        public void SaveEmail(string subject, string from, string body, bool htmlBody)
        {
            try
            {
                LogIn();

                properties = new MessageProperties()
                {
                    Subject = subject,
                    Body = body,
                    HtmlBody = htmlBody,
                    From = from
                };
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed to save email settings.\n" + e, "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            viewModel.EmailConfigured = true;
            if (viewModel.EmailConfigured && viewModel.CertificateGenerated)
                viewModel.SendEmailsButtonActive = true;
        }

        public void LogIn()
        {
            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "tokens";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            service = new GmailService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

        }

        public void SendEmails(List<Person> people)
        {
            BackgroundWorker bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true;
            bw.DoWork += SendEmail;
            bw.ProgressChanged += SendEmailProgressChanged;
            bw.RunWorkerCompleted += SendEmailComplete;
            bw.RunWorkerAsync(people);
        }

        private void SendEmail(object sender, DoWorkEventArgs args)
        {
            viewModel.ExcelButtonActive = false;
            viewModel.PdfButtonActive = false;
            viewModel.GenerateButtonActive = false;
            viewModel.EmailSettingsButtonActive = false;
            viewModel.SendEmailsButtonActive = false;

            var people = args.Argument as List<Person>;

            try
            {
                megaClient.Login("anna.kuznetsova130@gmail.com", "catherine1987!");

                IEnumerable<INode> megaNodes = megaClient.GetNodes();
                INode root = megaNodes.Single(x => x.Type == NodeType.Root);
                
                INode certsFolder = megaNodes.SingleOrDefault(x => x.ParentId == root.Id && x.Type == NodeType.Directory && x.Name == properties.Subject) ?? megaClient.CreateFolder(properties.Subject, root);

                List<string> attachmentLinks = new List<string>();

                INode attachmentsFolder = megaNodes.SingleOrDefault(x => x.ParentId == certsFolder.Id && x.Type == NodeType.Directory && x.Name == "Attachments") ?? megaClient.CreateFolder("Attachments", certsFolder);
                foreach (string file in attachments)
                {
                    INode uploadedFile = megaClient.UploadFile(file, attachmentsFolder);
                    string link = megaClient.GetDownloadLink(uploadedFile).AbsoluteUri;
                    attachmentLinks.Add(link);
                }

                for (int i = 0; i < people.Count; i++)
                    try
                    {
                        Person person = people[i];
                        BodyBuilder bodyBuilder = new BodyBuilder();

                        ProcessAttachments();

                        InternetAddress from = InternetAddress.Parse(credential.UserId);
                        from.Name = properties.From;

                        MimeMessage message = new MimeMessage(new[] { from }, new[] { InternetAddress.Parse(person.Email) },
                            properties.Subject,
                            bodyBuilder.ToMessageBody());

                        service.Users.Messages.Send(MessageFromMime(message), "me").Execute();

                        void ProcessAttachments()
                        {
                            string userFolderName = $"{person.LastName} {person.Name} {person.Patronymic} {viewModel.FirstCertificate + i}";
                            INode userFolder = megaNodes.SingleOrDefault(x => x.ParentId == certsFolder.Id && x.Type == NodeType.Directory && x.Name == userFolderName) ?? megaClient.CreateFolder(userFolderName, certsFolder);

                            List<string> fileLinks = new List<string>();
                            foreach (string file in person.Certificates)
                            {
                                INode uploadedFile = megaClient.UploadFile(file, userFolder);
                                string link = megaClient.GetDownloadLink(uploadedFile).AbsoluteUri;
                                fileLinks.Add(link);
                            }
                            fileLinks.AddRange(attachmentLinks);

                            string toAdd = string.Empty;
                            if (properties.HtmlBody)
                            {
                                foreach (string link in fileLinks)
                                    toAdd += $"<a href=\"{link}\">{link}</a><br><br>";
                                bodyBuilder.HtmlBody = properties.Body != null ? properties.Body.Replace("{links}", toAdd) : "";
                            }
                            else
                            {
                                foreach (string link in fileLinks)
                                    toAdd += $"{link}\n\n";
                                bodyBuilder.HtmlBody = properties.Body != null ? properties.Body.Replace("{links}", toAdd) : "";
                            }
                        }
                        Message MessageFromMime(MimeMessage msg)
                        {
                            ByteArrayOutputStream stream = new ByteArrayOutputStream();
                            msg.WriteTo(stream);
                            byte[] bytes = stream.ToArray();
                            string raw = Convert.ToBase64String(bytes);
                            string rawUri = new StringBuilder(raw)
                                .Replace('+', '-')
                                .Replace('/', '_').ToString();
                            return new Message { Raw = rawUri };
                        }
                    }
                    catch (Exception e)
                    {
                        if (e.GetType() == typeof(ApiException) && e.Message == "API response: ResourceNotExists")
                        {
                            i--;
                            continue;
                        }
                        var result = MessageBox.Show($"Failed to send message {i + 1}.\n{e}", "Error",
                            MessageBoxButtons.AbortRetryIgnore,
                            MessageBoxIcon.Error);
                        if (result == DialogResult.Retry)
                            i--;
                        else if (result == DialogResult.Abort)
                            break;
                    }
                    finally
                    {
                        ((BackgroundWorker)sender).ReportProgress(i * 100 / people.Count);
                    }
            }
            catch (Exception e)
            {
                MessageBox.Show($"Failed to send messages.\n{e}", "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                megaClient.Logout();
            }
        }

        private void SendEmailProgressChanged(object sender, ProgressChangedEventArgs args)
        {
            viewModel.ProgressBarValue = args.ProgressPercentage;
        }

        private void SendEmailComplete(object sender, RunWorkerCompletedEventArgs args)
        {
            viewModel.PdfButtonActive = true;
            viewModel.GenerateButtonActive = true;
            viewModel.EmailSettingsButtonActive = true;
            viewModel.SendEmailsButtonActive = true;

            viewModel.ProgressBarValue = 0;

            ((BackgroundWorker)sender).Dispose();
        }

        public struct MessageProperties
        {
            public string Subject;
            public string From;
            public string Body;
            public bool HtmlBody;
        }
    }
}
