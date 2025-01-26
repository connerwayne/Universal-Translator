// This file is a C# code file.
// This file is a part of the 'AuditTool' project.
// This file is the code-behind for the 'MainWindow.xaml' file.

// Currently no send email process. Need to figure out how to pull from a database such as Workday to find users managers. Will look into Farpa code for guidance.

//  There is one button that performs all necessary steps for the user to send user audit approval emails.  It is designed to be a one-click process for the user, and auto-detects the incoming file format.
//  The system supports the following import file formats: XLSX, XLS, ODS, CSV

//  The ImportFileButton method is used to import the user list file, convert it to a DataTable, sort the data, and save it to an Excel file.
//  The ProcessFileButton_Click method is the clickable button used to load the sorted data from the Excel file (ImportFileButtonh, convert the "LastLoginDate" to a standard DateTime format,
//  sort the data by the converted "LastLoginDate", and save the sorted data to a new Excel file.

//  This is only for testing purposes. The final version will be a console application that will run on a server and process the files automatically.
//  Maybe an automatic import from email attachments. User sends an email with the vendors user list attached, the system processes the file and sends the approval emails to the users managers and
//  BankSystems/IAM.
//
//  This would eliminate the need for the user to manually import the files and process them, saving time and reducing the risk of errors.
//  Furthermore, this program could be expanded to maintain a list of email responses with approval/denials and non-respones. This could also allow the system to automatically send follow-up emails to
//  users who have not responded, and eliminhates the manual effort required to create these records. This would require a database to store the email responses and a way to track the status of each user.

//  Ultimately, something like an MS Teams Power App with a really nice GUI would be ideal here. From a security perspective, the system would rely on SSO to authenticate users. Integration with MS tools 
//  would be seamless and would allow for easy tracking of user responses and approvals. The system could also be integrated with IAM systems (SailPoint?) to automatically update user access based on the responses, in theory.
//  I am unsure on the backend methods needed to store the user responses and track the status of each user. I would need to research the best way to implement this feature.
//  I would also need to research the best way to integrate with IAM systems to automatically update user access based on the responses. This would likely require an API integration.


//  ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//  The purpose of this tool is to take various vendors user access lists and convert them to a single standard format with the same header names for like data.
//  The tool should take your user list file and convert the header names to standard names in the output file. If it doesn't, you will need to add the new header names to the dictionary.
//  This tool is designed to work with both .csv and .xlsx files. If you have an Excel file, it will be converted to a .csv file before processing. This code is 
//  for a WPF application that allows the user to import a user list file, process the file, and generate a report. The report is saved to an Excel file with the data sorted by the "LastLoginDate" column.

//  The dictionary (located in the LoadCsvIntoDataTable method) is the key to the system. It takes the incoming header names from the various vendor files and converts them
//  to a standard header name used in the output file using a collection of key-value pairs. This keeps the data consistent throughout the process.
//  In the dictionary section, the key is the incoming header name from the user list file. The value is the standard header name used in the output file.
//  For example, the incoming header name is "User Name" and the standard header name is "UserName".
//  The dictionary is not case-sensitive, so it will work with any case of the incoming header name.

//  The CsvReader object is used to read the CSV file and convert it to a DataTable. The DataTable is then sorted using the SortDataTable method.
//  The sorted DataTable is saved to an Excel file using the SaveDataTableToExcel method.
//  The LoadDataFromExcel method is used to load the data from the Excel file into a DataTable.
//  The LoadCsvIntoDataTable method reads the CSV file and converts it to a DataTable. The DataTable is then sorted using the SortDataTable method.
//  The sorted DataTable is saved to an Excel file using the SaveDataTableToExcel method.

//  Public class User is used to define the properties of a user object. The properties are mapped to the standard header names used in the output file.
//  It is referenced in 3 methods, LoadUsersFromCsv, ConstructEmailBody, and SendEmail. This is used to load the user data from the CSV file, construct
//  the email body, and send the approval emails to the users managers.


using CsvHelper;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Outlook = Microsoft.Office.Interop.Outlook;
//using System.Text;
using System.Diagnostics.Eventing.Reader;
//using System.Printing;
using System.Collections.Specialized;
using ExcelDataReader;

namespace AuditTool
{
    public partial class MainWindow : Window
    {
        private DataTable LoadCsvIntoDataTable(string csvFilePath)
        {            
            var headerMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {       
            // THIS IS THE DICTIONARY SECTION.
            // EDIT TO ADD NEW HEADER NAMES FROM VENDORS ACTIVE USER REPORTS.
            // THIS DICTIONARY TRANSLATES THE VARIOUS INCOMING HEARDER NAMES INTO STANDARD HEADER NAMES.
            // { "Incoming" , "Outgoing" }   -   conversion mapping into C:\Temp\AuditTool\Output\SortedData.xlsx -
            // This takes a string input from the ingest document and converts it to the static output header
            // in the xls sheet.This keeps the datapoints consistent throughout the rest of the data manipulation processes.
            // Basically it takes possible header names for like data and renames the output header to a standard name used
            // in the SortedData.xls sheet.
            // If the header name is not in the Dictionary, the incoming header name will carry over to the output file.

                // input header -> UserEmail
                { "UserName", "UserEmail" }, //RetailWeb Email Address
                { "EMail", "UserEmail" },
                { "E-Mail", "UserEmail" },
                { "E Mail", "UserEmail" },
                { "Email Address", "UserEmail" },
                { "Employee Email", "UserEmail" }, //BlackBaud
                
                
                // input header -> HrJobTitle
                { "Job Title", "HrJobTitle" },
                { "HR Job Title", "HrJobTitle" },
                { "HRJobTitle", "HrJobTitle" },                     

                // input header -> UserName
                { "Users Name", "UserName" },
                { "User Name", "UserName" },
                { "User ID", "UserName" },
                { "Company/User Name", "UserName" },
                { "Login", "UserName" },
                { "Login ", "UserName" }, //Compass, added space
                { "Fiserv ID", "UserName" }, //CheckFree PartnerCare
                
                // input header -> UserFullName
                { "FirstName" + " " + "LastName", "UserFullName" }, //DocuSign - not working to combine. no header name being added to datatable. Need if merge function?
                { "First name" + " " + "Last name", "UserFullName" },
                { "Employee First name" + " " + "Employee Last name", "UserFullName" }, //Blackbaud - not working to combine. no header name being added to datatable.
                { "Display Name", "UserFullName" },
                { "Name4 ", "UserFullName" }, //Compass
                { "Rep Name", "UserFullName" }, //CheckFree PartnerCare
                             
                // input header -> LastLoginDate
                { "Last Login", "LastLoginDate" },
                { "Last Login ", "LastLoginDate" }, //compass
                { "Last Log In", "LastLoginDate" },
                { "Last Login Date (Avail 3.2)", "LastLoginDate" },
                { "LastLogin", "LastLoginDate" },
                { "Last_Login", "LastLoginDate" },
                { "Last date of login?", "LastLoginDate" },
                { "Last date of logon?", "LastLoginDate" },
                { "Last login Date?", "LastLoginDate" },
                { "Last Login Date", "LastLoginDate" },
                { "LastLoginDate", "LastLoginDate" },
                { "Last_Login_Date", "LastLoginDate" },

                // input header -> IsActive = bool True/False or yes/no?
                { "isSso", "IsActive" },
                { "Active Account?", "IsActive" },
                { "ActiveAccount", "IsActive" },
                { "User Active?", "IsActive" },
                { "User Status", "IsActive" },
                { "Active", "IsActive" },
                // input header -> ActiveOrInactive = bool Active/Inactive
                { "STATUS", "AvtiveOrInactive" }, //RetailWeb   //CheckFree Partnercare        
                                       
                // input header -> IsAdmin
                { "Admin", "IsAsmin" },
                { "CustomerAdmin", "IsAdmin" },      //RetailWeb
                { "Group", "IsAdmin" },      //DocuSign - Everyone = no, Administrators = yes
                
                
                //---------------------------------------------------------------------------------
                //User Access Role types. Theere are 3 fields available for various roles fields. 
                
                // input header -> UserRoles1
                { "Role", "UserRoles1" }, //IronMountain
                { "User Type", "UserRoles1" }, //Weiland
                { "UserGroup", "UserRoles1" }, //RetailWeb
                { "RoleName", "UserRoles1" }, //RemitWeb                
                { "Business Unit", "UserRoles1" }, //RemitWeb
                { "Role Name", "UserRoles1" }, //BlackBaud
                { "eSignPermissionProfile", "UserRoles1" }, //DocuSign
                { "jobTitle", "UserRoles1" }, //Grants Connect
                { "Role ", "UserRoles1" }, //Compass
                { "Assigned Role", "UserRoles1" }, //CheckFree PartnerCare
                
                // input header -> UserRoles2
                { "Line of Business", "UserRoles2" }, //IronMountain Connect
                { "Exception Management Role", "UserRoles2" }, //Weiland
                { "Services", "UserRoles2" },               
                
                // input header -> UserRoles3
                { "Reports", "UserRoles3" }, //IronMountain Connect
                { "Alternate Approver", "UserRoles3" }, // Weiland 
                { "Payment Sub-permissions", "UserRoles3" }
                
                //-----------------------------------------------------------------------------------   
            };
            DataTable dataTable = new DataTable();
            using (var reader = new StreamReader(csvFilePath))
            {
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    // Read the CSV header
                    csv.Read();
                    csv.ReadHeader();
                    foreach (string header in csv.HeaderRecord)
                    {
                        if (string.IsNullOrWhiteSpace(header))
                        {
                            string errorMessage = "CSV file contains an empty header.";
                            LogErrorAndSendEmail(errorMessage);
                            throw new ArgumentException(errorMessage);
                        }
                        string columnName = headerMapping.ContainsKey(header) ? headerMapping[header] : header;
                        if (!dataTable.Columns.Contains(columnName))
                        {
                            dataTable.Columns.Add(columnName);
                        }
                    }
                    // Read the CSV data
                    while (csv.Read())
                    {
                        var row = dataTable.NewRow();
                        foreach (string header in csv.HeaderRecord)
                        {
                            string columnName = headerMapping.ContainsKey(header) ? headerMapping[header] : header;
                            if (dataTable.Columns.Contains(columnName))
                            {
                                row[columnName] = csv.GetField(header);
                            }
                            else
                            {
                                throw new ArgumentException($"Column '{columnName}' does not belong to table.");
                            }
                        }
                        dataTable.Rows.Add(row);
                    }
                }
            }
            return dataTable;
        }
        public class User
        {
            public string IsActive { get; set; }
            public string UserName { get; set; }
            public string UserEmail { get; set; }
            public string HrRole { get; set; }
            public string UserRoles1 { get; set; }
            public string UserRoles2 { get; set; }
            public string UserRoles3 { get; set; }
            public DateTime LastLoginDate { get; set; }
            public string UsersManager { get; set; }
            public string ManagerEmail { get; set; }
        }
    //  private void ImportFileButton_Click(object sender, RoutedEventArgs e)
        private void ProcessFileButton_Click(object sender, RoutedEventArgs e)
        {
            // Open a file dialog to select the user list file
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "All files (*.*)|*.*|.csv files (*.csv)|*.csv|.xlsx files (*.xlsx)|*.xls*"; // Add more file types if needed. // First in list is default
            openFileDialog.InitialDirectory = @"C:\Temp\AuditTool\Ingest Files\test";
            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                string fileExtension = Path.GetExtension(filePath).ToLower();
                if (fileExtension == ".xls" || fileExtension == ".xlsx")
                {
                    filePath = ConvertExcelToCsv(filePath);
                    fileExtension = ".csv";
                }
                if (fileExtension != ".csv")
                {
                    MessageBox.Show("Unsupported file format.", "Error");
                    return;
                }
                DataTable dataTable = LoadCsvIntoDataTable(filePath);
                if (dataTable == null)
                {
                    MessageBox.Show("Failed to load data from the file.", "Error");
                    return;
                }

                DataTable sortedDataTable = SortDataTable(dataTable);
                string targetDirectory = @"C:\Temp\AuditTool\Output";

                // Ensure the target directory exists
                if (!Directory.Exists(targetDirectory))
                {
                    Directory.CreateDirectory(targetDirectory);
                }

                string excelFilePath = Path.Combine(targetDirectory, "SortedData.xlsx");
                SaveDataTableToExcel(sortedDataTable, excelFilePath);
                MessageBox.Show($"File has been processed and saved to {excelFilePath}", "File Processed");
            }
            // Process the file and generate the report
            {
                string sortedDataFilePath = "C:\\Temp\\AuditTool\\Output\\SortedData.xlsx";
                string reportFilePath = "C:\\Temp\\AuditTool\\Output\\ReportData.xlsx";
                // Load the data from the SortedData.xlsx file
                DataTable dataTable = LoadDataFromExcel(sortedDataFilePath);
                // Check if the "LastLoginDate" column exists
                if (dataTable.Columns.Contains("LastLoginDate"))
                {
                    // Add a new column for the converted DateTime values
                    dataTable.Columns.Add("ConvertedLastLoginDate", typeof(DateTime));
                    // Iterate through each row and convert the "LastLoginDate" to a standard DateTime format
                    foreach (DataRow row in dataTable.Rows)
                    {
                        if (DateTime.TryParse(row["LastLoginDate"].ToString(), out DateTime parsedDate))
                        {
                            row["ConvertedLastLoginDate"] = parsedDate;
                        }
                        else
                        {
                            // Handle invalid date format, e.g., set to a default value or remove the row
                            row["ConvertedLastLoginDate"] = DateTime.MinValue; // Example: setting to DateTime.MinValue
                        }
                    }
                    // Sort the data by the converted "LastLoginDate"
                    var sortedData = dataTable.AsEnumerable()
                        .OrderBy(row => row.Field<DateTime>("ConvertedLastLoginDate"))
                        .CopyToDataTable();
                    // Remove the temporary "ConvertedLastLoginDate" column
                    sortedData.Columns.Remove("ConvertedLastLoginDate");
                    // Save the sorted data to a new Excel file
                    SaveDataToExcel(sortedData, reportFilePath);
                    MessageBox.Show("Report generated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    // Continue processing without the "LastLoginDate" column
                    SaveDataToExcel(dataTable, reportFilePath);
                    MessageBox.Show("Report generated successfully without 'LastLoginDate' sorting!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            } 
        }

        private string ConvertExcelToCsv(string excelFilePath)
        {
            // System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string csvFilePath = Path.ChangeExtension(excelFilePath, ".csv");
            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports: XLSX, XLS, ODS, CSV
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    var dataTable = result.Tables.Cast<DataTable>().FirstOrDefault();
                    if (dataTable != null)
                    {
                        using (var writer = new StreamWriter(csvFilePath))
                        using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                        {
                            // Write the header
                            foreach (DataColumn column in dataTable.Columns)
                            {
                                csv.WriteField(column.ColumnName);
                            }
                            csv.NextRecord();
                            // Write the rows
                            foreach (DataRow row in dataTable.Rows)
                            {
                                foreach (DataColumn column in dataTable.Columns)
                                {
                                    csv.WriteField(row[column]);
                                }
                                csv.NextRecord();
                            }
                        }
                    }
                }
            }
            return csvFilePath;
        }
        private DataTable LoadDataFromExcel(string filePath)
        {
            DataTable dataTable = new DataTable();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                bool hasHeader = true;
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                    DataRow row = dataTable.NewRow();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                    dataTable.Rows.Add(row);
                }
            }
            return dataTable;
        }
        private void SaveDataToExcel(DataTable dataTable, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("ReportData");
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                    }
                }
                package.SaveAs(new FileInfo(filePath));
            }
        }
        private void LogErrorAndSendEmail(string errorMessage)
        {
            // Log the error  
            // (you can change the logging mechanism here)
            Console.WriteLine(errorMessage);
            // Create and send an email using Outlook
            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = "Universal Translator - Processing Error";
                mailItem.To = "ConnerConnects@gmail.com";
                mailItem.Body = $"An error occurred while processing the ingest file:\n\n{errorMessage}";
                mailItem.Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"LogErrorAndSendEmail method - Failed to send email: {ex.Message}");
            }
        }
        private DataTable SortDataTable(DataTable dataTable)
        {
            DataView dataView = dataTable.DefaultView;
            return dataView.ToTable();
        }
        private void SaveDataTableToExcel(DataTable dataTable, string excelFilePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("SortedData");
                worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
                package.SaveAs(new FileInfo(excelFilePath));
            }
        }
        private void LoadUsersFromCsv(string filePath)
        {
            var users = new List<User>();
            // Define the dictionary to map alternative header names to standard header names
            var headerMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
            };
            try
            {
                var lines = File.ReadAllLines(filePath);
                var headers = lines[0].Split(','); // Read the header line directly from the file
                foreach (var line in lines.Skip(1)) // Skip header line
                {
                    var columns = line.Split(',');
                    var user = new User();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        var header = headers[i].Trim();
                        if (headerMapping.TryGetValue(header, out var propertyName))
                        {
                            var property = typeof(User).GetProperty(propertyName);
                            if (property != null)
                            {
                                var value = columns[i].Trim();
                                if (property.PropertyType == typeof(DateTime))
                                {
                                    property.SetValue(user, DateTime.ParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture));
                                }
                                else if (property.PropertyType == typeof(bool))
                                {
                                    property.SetValue(user, bool.Parse(value));
                                }
                                else
                                {
                                    property.SetValue(user, Convert.ChangeType(value, property.PropertyType));
                                }
                            }
                        }
                    }
                    users.Add(user);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading ingest file for LoadUsersFromCsv method: {ex.Message}");
            }          
        }
        private DataTable LoadExcelIntoDataTable(string excelFilePath)
        {
            DataTable dataTable = new DataTable();
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                bool hasHeader = true; // adjust it accordingly
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");
                }
                var startRow = hasHeader ? 2 : 1;
                for (var rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                    DataRow row = dataTable.NewRow();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                    dataTable.Rows.Add(row);
                }
            }
            return dataTable;
        }

        //////////// this processes the file in one step. simple csv to xls conversion. 
        //////////private void OnButtonClick(object sender, RoutedEventArgs e)
        //////////{
        //////////    var csvFilePath = "path_to_your_csv_file.csv"; // Replace with your CSV file path
        //////////    var excelFilePath = "path_to_your_excel_file.xlsx"; // Replace with your desired Excel file path
        //////////                                                        // ParseCsvToExcel(csvFilePath, excelFilePath);
        //////////    MessageBox.Show(".CSV input file has been processed and saved to Excel format (.XLS) successfully!");
        //////////}


        //////////////////private void ManagerEmailsCheckBox_Checked(object sender, RoutedEventArgs e)
        //////////////////{
        //////////////////    string excelFilePath = @"C:\Temp\AuditTool\Output\SortedData.xlsx";
        //////////////////    DataTable dataTable = LoadExcelIntoDataTable(excelFilePath);
        //////////////////    foreach (DataRow row in dataTable.Rows)
        //////////////////    {
        //////////////////        string userEmail = row["UserEmail"].ToString();
        //////////////////        string emailBody = ConstructEmailBody(row);
        //////////////////        SendEmail(userEmail, emailBody);
        //////////////////    }
        //////////////////}



        //////////////////private string ConstructEmailBody(DataRow row)
        //////////////////{
        //////////////////    StringBuilder emailBody = new StringBuilder();
        //////////////////    emailBody.Append("<html><body><table border='1'>");
        //////////////////    foreach (DataColumn column in row.Table.Columns)
        //////////////////    {
        //////////////////        emailBody.AppendFormat("<tr><td><b>{0}</b></td><td>{1}</td></tr>", column.ColumnName, row[column]);
        //////////////////    }
        //////////////////    emailBody.Append("</table></body></html>");
        //////////////////    return emailBody.ToString();
        //////////////////}
        //////////////////private void SendEmail(string userEmail, string emailBody)
        //////////////////{
        //////////////////    Outlook.Application outlookApp = new Outlook.Application();
        //////////////////    Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
        //////////////////    mailItem.Subject = "Your Data from SortedData.xlsx";
        //////////////////    mailItem.To = userEmail;
        //////////////////    mailItem.HTMLBody = emailBody;
        //////////////////    mailItem.Send();
        //////////////////}   
    }    
}

//button template for the xaml file
// <Button Content="ImportFileButton" Click="ImportFileButton_Click" Template="{StaticResource CustomButtonTemplate}" Width="150" Height="50" Margin="22,75,350,200" />