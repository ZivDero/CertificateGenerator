using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows;
using CertificateGenerator.Other;
using CertificateGenerator.ViewModel;
using iText.IO.Font;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Element;
using PdfiumViewer;
using FontStyles = iText.IO.Font.Constants.FontStyles;
using FontWeights = System.Windows.FontWeights;
using ImageSource = System.Windows.Media.ImageSource;
using PdfDocument = iText.Kernel.Pdf.PdfDocument;
using TextAlignment = iText.Layout.Properties.TextAlignment;


namespace CertificateGenerator.Handlers
{
    public class PdfHandler
    {
        public FileInfo File;
        private readonly MainViewModel viewModel;

        private PdfiumViewer.PdfDocument previewDocument;

        public bool LoadFile(string fileName)
        {
            File = new FileInfo(fileName);

            try
            {
                previewDocument = PdfiumViewer.PdfDocument.Load(File.FullName);
            }
            catch
            {
                return false;
            }

            return true;
        }

        public void GenerateCertificates(string path, MainViewModel.LabelParameters nameParameters,
            MainViewModel.TextAlignment nameAlignment, MainViewModel.LabelParameters certParameters,
            MainViewModel.TextAlignment certAlignment, int certNumber, int dpi, bool addNumber, bool addZeroes, int digitCount)
        {
            BackgroundWorker bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true;
            bw.DoWork += Generate;
            bw.ProgressChanged += GenerateProgressChanged;
            bw.RunWorkerCompleted += GenerateComplete;

            List<object> args = new List<object>
            {
                path, nameParameters, nameAlignment, certParameters, certAlignment, certNumber, dpi, addNumber, addZeroes,
                digitCount
            };
            bw.RunWorkerAsync(args);
        }

        private void Generate(object sender, DoWorkEventArgs args)
        {
            try
            {
                viewModel.ExcelButtonActive = false;
                viewModel.PdfButtonActive = false;
                viewModel.GenerateButtonActive = false;
                viewModel.EmailSettingsButtonActive = false;
                viewModel.SendEmailsButtonActive = false;

                List<object> argsList = args.Argument as List<object>;

                string path = (string) argsList[0];
                MainViewModel.LabelParameters nameParameters = (MainViewModel.LabelParameters) argsList[1];
                MainViewModel.TextAlignment nameAlignment = (MainViewModel.TextAlignment) argsList[2];
                MainViewModel.LabelParameters certParameters = (MainViewModel.LabelParameters) argsList[3];
                MainViewModel.TextAlignment certAlignment = (MainViewModel.TextAlignment) argsList[4];
                int certNumber = (int) argsList[5];
                int dpi = (int) argsList[6];
                bool addNumber  = (bool)argsList[7];
                bool addZeroes = (bool) argsList[8];
                int digitCount = (int) argsList[9];

                FontProgramFactory.RegisterSystemFontDirectories();

                int nameFontStyle = GetFontStyle(nameParameters.Font),
                    certFontStyle = GetFontStyle(certParameters.Font);
                FontProgram nameFontProgram =
                        FontProgramFactory.CreateRegisteredFont(nameParameters.Font.Family.Source, nameFontStyle),
                    certFontProgram =
                        FontProgramFactory.CreateRegisteredFont(certParameters.Font.Family.Source, certFontStyle);

                float dpiRatio = (float) 72 / dpi;

                DeviceRgb nameColorDevice =
                        new DeviceRgb(nameParameters.Color.R, nameParameters.Color.G, nameParameters.Color.B),
                    certColorDevice =
                        new DeviceRgb(certParameters.Color.R, certParameters.Color.G, certParameters.Color.B);

                Color nameColor = Color.MakeColor(nameColorDevice.GetColorSpace(), nameColorDevice.GetColorValue());
                Color certColor = Color.MakeColor(certColorDevice.GetColorSpace(), certColorDevice.GetColorValue());

                int currentCert = certNumber;
                int personNumber = 1;
                int personCount = viewModel.ExcelHandler.People.Count;
                bool yesToAll = false;

                foreach (var person in viewModel.ExcelHandler.People)
                {
                    args.Result = new List<object> {path, certNumber};

                    string newPath = $"{path}\\{person.LastName} {person.Name} {person.Patronymic} {currentCert} ({person.Certificates.Count + 1}).pdf";

                    if (System.IO.File.Exists(newPath))
                    {
                        if (!yesToAll)
                        {
                            var response = MessageBox.Show("Certificate already exists. Use existing certificate?\nYes = yes for all, No = yes, Cancel = No", "Certificate exists", MessageBoxButton.YesNoCancel,
                                MessageBoxImage.Question);
                            if (response == MessageBoxResult.Yes)
                            {
                                yesToAll = true;
                                NextCert();
                                continue;
                            }

                            if (response == MessageBoxResult.No)
                            {
                                NextCert();
                                continue;
                            }
                        }
                        else
                        {
                            NextCert();
                            continue;
                        }
                    }

                    void NextCert()
                    {
                        person.Certificates.Add(newPath);
                        (sender as BackgroundWorker).ReportProgress(personNumber * 100 / personCount);
                        personNumber++;
                        currentCert++;
                    }

                    PdfReader reader = new PdfReader(File);
                    PdfWriter writer = new PdfWriter(newPath);
                    PdfDocument pdf = new PdfDocument(reader, writer);
                    PdfPage page = pdf.GetFirstPage();
                    PdfCanvas pdfCanvas = new PdfCanvas(page);


                    PdfFont namePdfFont = PdfFontFactory.CreateFont(nameFontProgram, "Cp1251", true),
                        certPdfFont = PdfFontFactory.CreateFont(certFontProgram, "Cp1251", true);

                    Canvas canvas = new Canvas(pdfCanvas, page.GetPageSizeWithRotation());

                    float x,
                        y = page.GetPageSizeWithRotation().GetHeight() -
                            (float) (nameParameters.Position.Y + nameParameters.Height / 2 +
                                     nameParameters.Font.Size / 2) *
                            dpiRatio;

                    // Name start
                    TextAlignment pdfNameAlignment;
                    if (nameAlignment == MainViewModel.TextAlignment.Left)
                    {
                        x = (float) nameParameters.Position.X * dpiRatio;
                        pdfNameAlignment = TextAlignment.LEFT;
                    }
                    else if (nameAlignment == MainViewModel.TextAlignment.Center)
                    {
                        x = (float) (nameParameters.Position.X + nameParameters.Width / 2) * dpiRatio;
                        pdfNameAlignment = TextAlignment.CENTER;
                    }
                    else
                    {
                        x = (float) (nameParameters.Position.X + nameParameters.Width) * dpiRatio;
                        pdfNameAlignment = TextAlignment.RIGHT;
                    }

                    Paragraph paragraph = new Paragraph()
                        .SetFont(namePdfFont)
                        .SetFontSize((int) nameParameters.Font.Size * dpiRatio)
                        .SetFontColor(nameColor)
                        .Add($"{person.LastName} {person.Name} {person.Patronymic}");

                    canvas.ShowTextAligned(paragraph, x, y, pdfNameAlignment);
                    // Name end

                    // Number start
                    if (addNumber)
                    {
                        TextAlignment pdfCertAlignment;
                        if (certAlignment == MainViewModel.TextAlignment.Left)
                        {
                            x = (float)certParameters.Position.X * dpiRatio;
                            pdfCertAlignment = TextAlignment.LEFT;
                        }
                        else if (certAlignment == MainViewModel.TextAlignment.Center)
                        {
                            x = (float)(certParameters.Position.X + certParameters.Width / 2) * dpiRatio;
                            pdfCertAlignment = TextAlignment.CENTER;
                        }
                        else
                        {
                            x = (float)(certParameters.Position.X + certParameters.Width) * dpiRatio;
                            pdfCertAlignment = TextAlignment.RIGHT;
                        }

                        y = page.GetPageSizeWithRotation().GetHeight() -
                            (float)(certParameters.Position.Y + certParameters.Height / 2 + certParameters.Font.Size / 2) *
                            dpiRatio;

                        string certText = currentCert.ToString();
                        if (addZeroes)
                            while (certText.Length < digitCount)
                                certText = "0" + certText;


                        paragraph = new Paragraph()
                            .SetFont(certPdfFont)
                            .SetFontSize((int)certParameters.Font.Size * dpiRatio)
                            .SetFontColor(certColor)
                            .Add(certText);

                        canvas.ShowTextAligned(paragraph, x, y, pdfCertAlignment);
                        // Number end
                    }

                    canvas.Close();
                    pdf.Close();
                    reader.Close();
                    writer.Close();

                    person.Certificates.Add(newPath);

                    (sender as BackgroundWorker).ReportProgress(personNumber * 100 / personCount);
                    personNumber++;
                    currentCert++;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed to generate certificates.\n" + e, "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void GenerateProgressChanged(object sender, ProgressChangedEventArgs args)
        {
            viewModel.ProgressBarValue = args.ProgressPercentage;
        }

        private void GenerateComplete(object sender, RunWorkerCompletedEventArgs args)
        {
            viewModel.PdfButtonActive = true;
            viewModel.GenerateButtonActive = true;
            viewModel.EmailSettingsButtonActive = true;

            viewModel.ProgressBarValue = 0;
            if (args.Result != null)
            {
                viewModel.CertificateGenerated = true;
                viewModel.ExcelHandler.GenerateList((string)(args.Result as List<object>)[0],
                    (int)(args.Result as List<object>)[1]);
            }
            
            if (viewModel.EmailConfigured && viewModel.CertificateGenerated)
                    viewModel.SendEmailsButtonActive = true;

            ((BackgroundWorker) sender).Dispose();
        }

        private int GetFontStyle(MainViewModel.Font font)
        {
            if (font.Weight == FontWeights.Bold && font.Style == System.Windows.FontStyles.Italic)
                return FontStyles.BOLDITALIC;
            if (font.Weight == FontWeights.Bold)
                return FontStyles.BOLD;
            if (font.Style == System.Windows.FontStyles.Italic)
                return FontStyles.ITALIC;
            return FontStyles.NORMAL;
        }

        public ImageSource GetPreview(int dpi = 96)
        {
            return previewDocument.Render(0, dpi, dpi, PdfRenderFlags.CorrectFromDpi).ToImageSource();
        }

        public PdfHandler(MainViewModel vm)
        {
            viewModel = vm;
        }
    }
}