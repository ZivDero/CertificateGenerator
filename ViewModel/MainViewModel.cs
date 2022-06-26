using CertificateGenerator.Handlers;
using CertificateGenerator.Properties;
using CertificateGenerator.Windows;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace CertificateGenerator.ViewModel
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public bool PdfLoaded;
        public bool ExcelLoaded = false;
        public bool EmailConfigured = false;
        public bool CertificateGenerated = false;

        public void InitializeViewModel()
        {
            PdfHandler = new PdfHandler(this);
            ExcelHandler = new ExcelHandler(this);
            SmtpHandler = new SmtpHandler(this);
        }

        public ExcelHandler ExcelHandler;
        public PdfHandler PdfHandler;
        public SmtpHandler SmtpHandler;

        private bool generateButtonActive;
        public bool GenerateButtonActive
        {
            get => generateButtonActive;
            set
            {
                generateButtonActive = value;
                OnPropertyChanged(nameof(GenerateButtonActive));
            }
        }

        private bool pdfButtonActive = true;
        public bool PdfButtonActive
        {
            get => pdfButtonActive;
            set
            {
                pdfButtonActive = value;
                OnPropertyChanged(nameof(PdfButtonActive));
            }
        }

        private bool excelButtonActive = true;
        public bool ExcelButtonActive
        {
            get => excelButtonActive;
            set
            {
                excelButtonActive = value;
                OnPropertyChanged(nameof(ExcelButtonActive));
            }
        }

        private bool emailSettingsButtonActive = true;
        public bool EmailSettingsButtonActive
        {
            get => emailSettingsButtonActive;
            set
            {
                emailSettingsButtonActive = value;
                OnPropertyChanged(nameof(EmailSettingsButtonActive));
            }
        }

        private bool sendEmailsButtonActive;
        public bool SendEmailsButtonActive
        {
            get => sendEmailsButtonActive;
            set
            {
                sendEmailsButtonActive = value;
                OnPropertyChanged(nameof(SendEmailsButtonActive));
            }
        }

        private int progressBarValue;
        public int ProgressBarValue
        {
            get => progressBarValue;
            set
            {
                progressBarValue = value;
                OnPropertyChanged(nameof(ProgressBarValue));
            }
        }

        private TextAlignment nameAlignment = TextAlignment.Left;
        public TextAlignment NameAlignment
        {
            get => nameAlignment;
            set
            {
                nameAlignment = value;
                OnPropertyChanged(nameof(NameAlignment));
            }
        }

        private TextAlignment numberAlignment = TextAlignment.Left;
        public TextAlignment NumberAlignment
        {
            get => numberAlignment;
            set
            {
                numberAlignment = value;
                OnPropertyChanged(nameof(NumberAlignment));
            }
        }

        private int dpi = 96;
        public int Dpi
        {
            get => dpi;
            set
            {
                dpi = value;
                OnPropertyChanged(nameof(Dpi));
            }
        }

        private int firstCertificate = 1;
        public int FirstCertificate
        {
            get => firstCertificate;
            set
            {
                firstCertificate = value;
                OnPropertyChanged(nameof(FirstCertificate));
            }
        }

        private bool addNumber = true;
        public bool AddNumber
        {
            get => addNumber;
            set
            {
                addNumber = value;
                OnPropertyChanged(nameof(AddNumber));
            }
        }

        private bool addZeroes;
        public bool AddZeroes
        {
            get => addZeroes;
            set
            {
                addZeroes = value;
                OnPropertyChanged(nameof(AddZeroes));
            }
        }

        private int digitCount = 5;
        public int DigitCount
        {
            get => digitCount;
            set
            {
                digitCount = value;
                OnPropertyChanged(nameof(NumberAlignment));
            }
        }

        private LabelParameters label1Parameters = new LabelParameters();
        public LabelParameters Label1Parameters
        {
            get => label1Parameters;
            set
            {
                label1Parameters = value;
                OnPropertyChanged(nameof(Label1Parameters));
            }
        }

        private LabelParameters label2Parameters = new LabelParameters();
        public LabelParameters Label2Parameters
        {
            get => label2Parameters;
            set
            {
                label2Parameters = value;
                OnPropertyChanged(nameof(Label2Parameters));
            }
        }

        private string nameTextPreview = "Ivanov Ivan Ivanovich";
        public string NameTextPreview
        {
            get => nameTextPreview;
            set
            {
                nameTextPreview = value;
                OnPropertyChanged(nameof(NameTextPreview));
            }
        }

        private string emailSubject = "Example";
        public string EmailSubject
        {
            get => emailSubject;
            set
            {
                emailSubject = value;
                OnPropertyChanged(nameof(EmailSubject));
            }
        }

        private string senderName = "Me";
        public string SenderName
        {
            get => senderName;
            set
            {
                senderName = value;
                OnPropertyChanged(nameof(SenderName));
            }
        }

        private bool htmlBody;
        public bool HtmlBody
        {
            get => htmlBody;
            set
            {
                htmlBody = value;
                OnPropertyChanged(nameof(HtmlBody));
            }
        }

        private string emailBody;
        public string EmailBody
        {
            get => emailBody;
            set
            {
                emailBody = value;
                OnPropertyChanged(nameof(EmailBody));
            }
        }

        private List<string> users;
        public List<string> Users
        {
            get => users;
            set
            {
                users = value;
                OnPropertyChanged(nameof(Users));
            }
        }

        private ImageSource imageSource;
        public ImageSource ImageSource
        {
            get => imageSource;
            set
            {
                imageSource = value;
                OnPropertyChanged(nameof(ImageSource));
            }
        }

        private RelayCommand openExcelCommand;
        public RelayCommand OpenExcelCommand
        {
            get
            {
                return openExcelCommand ??= new RelayCommand(obj =>
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.Filter = "Excel Files (*.xl;*.xlsx)|*.xl;*.xlsx|All Files (*.*)|*.*";
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == true)
                        ExcelHandler.LoadFile(openFileDialog.FileName);
                });
            }
        }

        private RelayCommand openPdfCommand;
        public RelayCommand OpenPdfCommand
        {
            get
            {
                return openPdfCommand ??= new RelayCommand(obj =>
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.Filter = "Adobe PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*";
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == true)
                    {
                        if (PdfHandler.LoadFile(openFileDialog.FileName))
                        {
                            ImageSource = PdfHandler.GetPreview(Dpi);
                            PdfLoaded = true;
                            if (PdfLoaded && ExcelLoaded)
                                GenerateButtonActive = true;
                        }
                    }
                });
            }
        }

        private RelayCommand selectFont1Command;
        public RelayCommand SelectFont1Command
        {
            get
            {
                return selectFont1Command ??= new RelayCommand(obj =>
                {
                    FontDialog dialog = new FontDialog();
                    if (dialog.ShowDialog() == DialogResult.OK)
                        Label1Parameters.Font = new Font(dialog, Dpi);
                });
            }
        }

        private RelayCommand selectFont2Command;
        public RelayCommand SelectFont2Command
        {
            get
            {
                return selectFont2Command ??= new RelayCommand(obj =>
                {
                    FontDialog dialog = new FontDialog();
                    if (dialog.ShowDialog() == DialogResult.OK)
                        Label2Parameters.Font = new Font(dialog, Dpi);
                });
            }
        }

        private RelayCommand generateCertificatesCommand;
        public RelayCommand GenerateCertificatesCommand
        {
            get
            {
                return generateCertificatesCommand ??= new RelayCommand(obj =>
                {
                    FolderBrowserDialog dialog = new FolderBrowserDialog {Description = Resources.SaveToFolderDesc};
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        PdfHandler.GenerateCertificates(dialog.SelectedPath, Label1Parameters, NameAlignment, Label2Parameters, NumberAlignment,
                            FirstCertificate, Dpi, AddNumber, AddZeroes, DigitCount);
                    }
                });
            }
        }

        private RelayCommand logInCommand;
        public RelayCommand LogInCommand
        {
            get
            {
                return logInCommand ??= new RelayCommand(obj =>
                {
                    SmtpHandler.SaveEmail(EmailSubject, SenderName, EmailBody, HtmlBody);
                });
            }
        }

        private RelayCommand saveEmailCommand;
        public RelayCommand SaveEmailCommand
        {
            get
            {
                return saveEmailCommand ??= new RelayCommand(obj =>
                {
                    SmtpHandler.SaveEmail(EmailSubject, SenderName, EmailBody, HtmlBody);
                });
            }
        }

        private RelayCommand addAttachmentCommand;
        public RelayCommand AddAttachmentCommand
        {
            get
            {
                return addAttachmentCommand ??= new RelayCommand(obj =>
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.Filter = "All Files (*.*)|*.*";
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == true)
                        SmtpHandler.AddAttachment(openFileDialog.FileName);
                });
            }
        }

        private RelayCommand emailSettingsCommand;
        public RelayCommand EmailSettingsCommand
        {
            get
            {
                return emailSettingsCommand ??= new RelayCommand(obj =>
                {
                    EmailWindow emailWindow = new EmailWindow();
                    emailWindow.Show();
                });
            }
        }

        private RelayCommand sendEmailsCommand;
        public RelayCommand SendEmailsCommand
        {
            get
            {
                return sendEmailsCommand ??= new RelayCommand(obj =>
                {
                    SmtpHandler.SendEmails(ExcelHandler.People);
                });
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public class Font : INotifyPropertyChanged
        {
            private FontFamily family;
            public FontFamily Family
            {
                get => family;
                set
                {
                    family = value;
                    OnPropertyChanged("Family");
                }
            }

            private double size;
            public double Size
            {
                get => size;
                set
                {
                    size = value;
                    OnPropertyChanged("Size");
                }
            }

            private FontWeight weight;
            public FontWeight Weight
            {
                get => weight;
                set
                {
                    weight = value;
                    OnPropertyChanged("Weight");
                }
            }

            private FontStyle style;
            public FontStyle Style
            {
                get => style;
                set
                {
                    style = value;
                    OnPropertyChanged("Style");
                }
            }

            public Font(FontDialog dialog, int dpi)
            {
                Family = new FontFamily(dialog.Font.Name);
                Size = dialog.Font.Size * dpi / 72.0;
                Weight = dialog.Font.Bold ? FontWeights.Bold : FontWeights.Regular;
                Style = dialog.Font.Italic ? FontStyles.Italic : FontStyles.Normal;
            }

            public Font()
            {
                Family = new FontFamily("Times New Roman");
                Size = 12;
                Weight = FontWeights.Normal;
                Style = FontStyles.Normal;
            }

            public event PropertyChangedEventHandler PropertyChanged;
            protected virtual void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public class LabelParameters : INotifyPropertyChanged
        {
            private Point position;
            public Point Position
            {
                get => position;
                set
                {
                    position = value;
                    OnPropertyChanged("Position");
                }
            }

            private double width;
            public double Width
            {
                get => width;
                set
                {
                    width = value;
                    OnPropertyChanged("Width");
                }
            }

            private double height;
            public double Height
            {
                get => height;
                set
                {
                    height = value;
                    OnPropertyChanged("Height");
                }
            }

            private Color color;
            public Color Color
            {
                get => color;
                set
                {
                    color = value;
                    OnPropertyChanged("Color");
                }
            }

            private Font font;
            public Font Font
            {
                get => font;
                set
                {
                    font = value;
                    OnPropertyChanged("Font");
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
            protected virtual void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }

            public LabelParameters()
            {
                position = new Point(0, 0);
                width = double.NaN;
                height = double.NaN;
                color = Colors.Black;
                font = new Font();
            }
        }

        public enum TextAlignment
        {
            Left, Center, Right
        }
    }
}
