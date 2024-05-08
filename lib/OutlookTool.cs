using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MnS.lib
{
    public static class OutlookTool
    {
        /// <summary>
        /// Create Outlook Email with an excel file attached <br></br> <br></br>
        /// Parameter: <br></br>
        /// "dataTable" Create excel file from data table
        /// </summary>
        /// <param name="dataTable"></param>
        public static void SendEmailWithExcelAttachment(DataTable dataTable)
        {
            try
            {
                string tempFilePath = Path.Combine(Path.GetTempPath(), "Test_Spare_Data.xlsx");

                ExcelTool.CreateExcelFileFromDataTable(dataTable, tempFilePath);

                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.MailItem mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

                Outlook.Attachment attachment = mailItem.Attachments.Add(tempFilePath);

                string emailBody = "<html><body>";
                emailBody += "<h3>Dear ... :</h3>";
                emailBody += "Email content...<br><br>";
                emailBody += "<br>Attached database:";
                emailBody += "<table border='1' cellspacing='0' cellpadding='5'>";
                emailBody += "<tr>";
                foreach (DataColumn col in dataTable.Columns)
                {
                    emailBody += $"<th>{col.ColumnName}</th>";
                }
                emailBody += "</tr>";

                foreach (DataRow row in dataTable.Rows)
                {
                    emailBody += "<tr>";
                    foreach (DataColumn col in dataTable.Columns)
                    {
                        emailBody += $"<td>{row[col]}</td>";
                    }
                    emailBody += "</tr>";
                }

                emailBody += "</table></body></html>";
                emailBody += "<br>Thanks and Best Regards,";

                mailItem.HTMLBody = emailBody;

                mailItem.Subject = "[MnS-Automail]: ---";
                mailItem.To = "---@vn.pepperl-fuchs.com";
                mailItem.CC = "---@vn.pepperl-fuchs.com";

                mailItem.Display();

                Marshal.ReleaseComObject(attachment);
                Marshal.ReleaseComObject(mailItem);
                Marshal.ReleaseComObject(outlookApp);

                File.Delete(tempFilePath);

                MessageBox.Show("Email sent successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error sending email: " + ex.Message);
            }
        }

        public static void Feedback(string receiver, string cc)
        {
            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.MailItem mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

                string emailBody = "<html><body>";
                emailBody += $"<h3>Hi {receiver} :</h3>";
                emailBody += "Email content...<br><br>";
                emailBody += "<tr>";
                emailBody += "</tr>";

                emailBody += "</table></body></html>";
                emailBody += "<br>Thanks and Best Regards,";

                mailItem.HTMLBody = emailBody;

                mailItem.Subject = "[MnS-Automail]: ---";
                mailItem.To = receiver + "@vn.pepperl-fuchs.com";
                mailItem.CC = cc + "@vn.pepperl-fuchs.com";

                mailItem.Display();
                Marshal.ReleaseComObject(mailItem);
                Marshal.ReleaseComObject(outlookApp);

                MessageBox.Show("Email created successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error sending email: " + ex.Message);
            }
        }
    }
}
