using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using CsvHelper;
using OfficeOpenXml;
using Microsoft.Win32;





namespace User_List_IO_GUI
{
    public partial class MainWindow : Window
    {
        //public MainWindow()
        //{
        //    InitializeComponent();
        //}
    private System.Windows.Controls.Button ImportFileButton;
    private System.Windows.Controls.Button ProcessButton;
    
    //private void InitializeComponent()
    //    {
    //        this.ImportFileButton = new System.Windows.Controls.Button();
    //        this.ImportFileButton.Content = "Import File";
    //        this.ImportFileButton.Click += new System.Windows.RoutedEventHandler(ImportFileButton_Click);

    //        this.ProcessButton = new System.Windows.Controls.Button();
    //        this.ProcessButton.Content = "Process CSV to Excel";
    //        this.ProcessButton.Click += new System.Windows.RoutedEventHandler(ProcessFileButton_Click);

    //        var stackPanel = new System.Windows.Controls.StackPanel();
    //        stackPanel.Children.Add(this.ImportFileButton);
    //        stackPanel.Children.Add(this.ProcessButton);

    //        Content = stackPanel;
    //    }

        private void ImportFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                // Read the file content
                string fileContent = File.ReadAllText(filePath);
                // Do something with the file content
                MessageBox.Show(fileContent, "File Content");
            }
        }
        private void ProcessFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                // Read the file content
                string fileContent = File.ReadAllText(filePath);
                // Do something with the file content
                MessageBox.Show(fileContent, "File Content");
            }
        }

        private void ParseCsvToExcel(string csvFilePath, string excelFilePath)
        {
            var users = new List<User>();
            using (var reader = new StreamReader(csvFilePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                users = csv.GetRecords
                    <User>().ToList();
            }
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Users");
                worksheet.Cells.LoadFromCollection(users, true);
                package.SaveAs(new FileInfo(excelFilePath));
            }
        }
        private void OnButtonClick(object sender, RoutedEventArgs e)
        {
            var csvFilePath = "path_to_your_csv_file.csv"; // Replace with your CSV file path
            var excelFilePath = "path_to_your_excel_file.xlsx"; // Replace with your desired Excel file path
            ParseCsvToExcel(csvFilePath, excelFilePath);
            MessageBox.Show(".CSV input file has been processed and saved to Excel format (.XLS) successfully!");
        }
    }
    public class User
    {
        public bool IsActive { get; set; }
        public string UserName { get; set; }
        public string Email { get; set; }
        public string HrRole { get; set; }
        public bool AccessRights { get; set; }
        public bool lastLogin { get; set; }
        public bool UsersManager { get; set; }
        public bool ManagerEmail { get; set; }
}
}