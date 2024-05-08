using System;
using System.Data;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MnS.lib
{
    public static class UserLogTool
    {
        public static readonly Outlook.Application outlookApp;
        public static DataTable userInfor;

        public static DataTable UserData(string userID, string fullName, string department, string email, string loginID, string rightAccess, string image)
        {
            DataTable userTable = new DataTable();
            userTable.Columns.Add("UserID", typeof(string));
            userTable.Columns.Add("FullName", typeof(string));
            userTable.Columns.Add("Department", typeof(string));
            userTable.Columns.Add("Email", typeof(string));
            userTable.Columns.Add("LoginID", typeof(string));
            userTable.Columns.Add("RightAccess", typeof(string));
            userTable.Columns.Add("Image", typeof(string));

            DataRow newRow = userTable.NewRow();
            newRow["UserID"] = userID;
            newRow["FullName"] = fullName;
            newRow["Department"] = department;
            newRow["Email"] = email;
            newRow["LoginID"] = loginID;
            newRow["RightAccess"] = rightAccess;
            newRow["Image"] = image;

            userTable.Rows.Add(newRow);
            userInfor = userTable;
            return userTable;
        }

        public static void UserLog(string action, string message)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(PathReader.Log_file, true))
                {
                    string logEntry = $"{DateTime.Now:dd/MM/yyyy HH:mm:ss tt} - Action: Account {userInfor.Rows[0][0]} {userInfor.Rows[0][1]} {GetInfor.Email_log} {action} - Message: {message}\n";
                    writer.Write(logEntry);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error writing to log file: " + ex.Message);
            }
        }

        public static void UserData(string message)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(PathReader.Log_data, true))
                {
                    string logEntry = $"{DateTime.Now:dd/MM/yyyy HH:mm:ss tt} - Account: {GetInfor.User_log}_{GetInfor.Dept_log}_{GetInfor.Email_log} - Action: {message}\n";
                    writer.Write(logEntry);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error writing to log file: " + ex.Message);
            }
        }
    }
}