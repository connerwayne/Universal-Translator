using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;

namespace User_List_IO_GUI
{
    public partial class MainWindow : Window
    {
        private void ProcessFileButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                InitialDirectory = @"C:\Users\Conne\OneDrive\Documents\1 Code\CSV IO Tool\User Lists"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                var filePath = openFileDialog.FileName;
                var fileContent = File.ReadAllLines(filePath);
                // Create a new Excel package
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package =
                       new ExcelPackage(new FileInfo(@"C:\TEMP\DEV\Response Verification Worksheet.xlsx")))
                {
                }

                using (var package = new ExcelPackage())

                {
                    // Add a new worksheet to the empty workbook
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    // Load the CSV data into the worksheet
                    for (var i = 0; i < fileContent.Length; i++)
                    {
                        var rowData = fileContent[i].Split(',');
                        for (var j = 0; j < rowData.Length; j++) worksheet.Cells[i + 1, j + 1].Value = rowData[j];
                    }

                    //// Save the Excel package to a file
                    //string excelFilePath = Path.Combine(@"C:\TEMP\DEV", "Response Verification Worksheet.xlsx");
                    //File.WriteAllBytes(excelFilePath, package.GetAsByteArray());
                    //package.Save();

                    // Prompt the user to choose where to save the Excel file
                    var saveFileDialog = new SaveFileDialog
                    {
                        Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                        DefaultExt = ".xlsx",
                        FileName = Path.ChangeExtension(Path.GetFileName(filePath), @"Response Verification Worksheet.xlsx")
                    };
                    if (saveFileDialog.ShowDialog() == true)
                    {
                        var excelFilePath = saveFileDialog.FileName;
                        File.WriteAllBytes(excelFilePath, package.GetAsByteArray());
                        MessageBox.Show("CSV data has been converted to Excel.", "Conversion Complete");
                        Environment.Exit(0);
                    }
                }
            }
        }

        public class User
        {
            public bool IsActive { get; set; }
            public string UserName { get; set; }
            public string Email { get; set; }
            public string HrRole { get; set; }
            public string AccessRights { get; set; }
            public string LastLogin { get; set; }
            public string UsersManager { get; set; }
            public string ManagerEmail { get; set; }
        }
    }
}