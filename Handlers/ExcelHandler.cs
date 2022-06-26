using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows;
using CertificateGenerator.ViewModel;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace CertificateGenerator.Handlers
{
    public class ExcelHandler
    {
        public FileInfo File;
        public List<Person> People;

        private readonly MainViewModel viewModel;

        public void LoadFile(string fileName)
        {
            BackgroundWorker bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true;
            bw.DoWork += Load;
            bw.ProgressChanged += LoadProgressChanged;
            bw.RunWorkerCompleted += LoadComplete;
            bw.RunWorkerAsync(fileName);
        }

        private void Load(object sender, DoWorkEventArgs args)
        {
            viewModel.ExcelButtonActive = false;
            viewModel.PdfButtonActive = false;
            viewModel.GenerateButtonActive = false;
            viewModel.EmailSettingsButtonActive = false;

            args.Result = true;

            string fileName = args.Argument as string;

            File = new FileInfo(fileName);
            People = new List<Person>();

            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Add();
            Worksheet worksheet = workbook.ActiveSheet;

            try
            {
                workbook = excel.Workbooks.Open(fileName);
                worksheet = workbook.Sheets[1];

                int rows = 0;

                while (!string.IsNullOrEmpty(GetCell(rows + 1, 1)))
                    rows++;

                for (int i = 1; i <= rows; i++)
                {
                    string lastName = GetCell(i, 1), name = GetCell(i, 2), patronymic = GetCell(i, 3), email = GetCell(i, 4);

                    while (true)
                    {
                        int index = lastName.IndexOfAny(Path.GetInvalidFileNameChars());
                        if (index != -1) lastName = lastName.Remove(index);
                        else break;
                    }

                    while (true)
                    {
                        int index = name.IndexOfAny(Path.GetInvalidFileNameChars());
                        if (index != -1) name = name.Remove(index);
                        else break;
                    }

                    while (true)
                    {
                        int index = patronymic.IndexOfAny(Path.GetInvalidFileNameChars());
                        if (index != -1) patronymic = patronymic.Remove(index);
                        else break;
                    }

                    People.Add(new Person(lastName, name, patronymic, email));
                    if (i < rows)
                        (sender as BackgroundWorker).ReportProgress(i * 100 / rows);
                }

                string GetCell(int row, int cell)
                {
                    var value = Convert.ToString(worksheet.Cells[row, cell].Value2);
                    if (value != null)
                        return value;
                    return "";
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed to load excel file.\n" + e, "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
                args.Result = false;
            }
            finally
            {
                workbook.Close();
                excel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                (sender as BackgroundWorker).ReportProgress(100);
            }
        }

        private void LoadProgressChanged(object sender, ProgressChangedEventArgs args)
        {
            viewModel.ProgressBarValue = args.ProgressPercentage;
        }

        private void LoadComplete(object sender, RunWorkerCompletedEventArgs args)
        {
            viewModel.ExcelButtonActive = true;
            viewModel.PdfButtonActive = true;
            viewModel.GenerateButtonActive = true;
            viewModel.EmailSettingsButtonActive = true;
            viewModel.ProgressBarValue = 0;

            if ((bool) args.Result)
                viewModel.ExcelLoaded = true;
            else
                viewModel.ExcelButtonActive = true;

            if (viewModel.PdfLoaded && viewModel.ExcelLoaded)
                viewModel.GenerateButtonActive = true;
        }

        public void GenerateList(string path, int firstCertificate)
        {
            BackgroundWorker bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true;
            bw.DoWork += Generate;
            bw.ProgressChanged += GenerateProgressChanged;
            bw.RunWorkerCompleted += GenerateComplete;

            List<object> args = new List<object> { path, firstCertificate };
            bw.RunWorkerAsync(args);
        }

        private void Generate(object sender, DoWorkEventArgs args)
        {
            viewModel.ExcelButtonActive = false;
            viewModel.PdfButtonActive = false;
            viewModel.GenerateButtonActive = false;
            viewModel.EmailSettingsButtonActive = false;

            args.Result = true;

            List<object> argsList = args.Argument as List<object>;

            string path = (string)argsList[0];
            int firstCertificate = (int)argsList[1];

            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Add();
            Worksheet worksheet = workbook.ActiveSheet;

            try
            {
                for (int i = 0; i < People.Count; i++)
                {
                    SetCell(i + 1, 1, People[i].LastName);
                    SetCell(i + 1, 2, People[i].Name);
                    SetCell(i + 1, 3, People[i].Patronymic);
                    SetCell(i + 1, 4, (i + firstCertificate).ToString());
                    ((BackgroundWorker) sender).ReportProgress(i * 100 / People.Count);
                }

                workbook.SaveAs($"{path}\\certificates.xlsx");
            }
            catch (Exception e)
            {
                args.Result = false;
                MessageBox.Show("Failed to save Excel file.\n" + e, "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                workbook.Close();
                excel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }

            void SetCell(int row, int cell, string value)
            {
                worksheet.Cells[row, cell].Value2 = value;
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

            if (viewModel.PdfLoaded && viewModel.ExcelLoaded)
                viewModel.GenerateButtonActive = true;

            ((BackgroundWorker)sender).Dispose();
        }

        public ExcelHandler(MainViewModel vm)
        {
            viewModel = vm;
        }
    }

    public struct Person
    {
        public string LastName;
        public string Name;
        public string Patronymic;
        public List<string> Certificates;
        public string Email;

        public Person(string lastName, string name, string patronymic, string email)
        {
            LastName = lastName;
            Name = name;
            Patronymic = patronymic;
            Certificates = new List<string>();
            Email = email;
        }
    }
}
