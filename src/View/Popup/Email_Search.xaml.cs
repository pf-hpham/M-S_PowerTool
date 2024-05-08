using System;
using MnS.lib;
using System.Data;
using System.Windows;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MnS
{
    public partial class Email_Search : Window
    {
        public Email_Search()
        {
            UserLogTool.UserData("Using Email search function");
            InitializeComponent();
            txtEmail.SelectAll();
            txtEmail.Focus();
        }

        private void Submit_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtEmail.Text.Trim()))
            {
                MessageBox.Show("Please enter your email address.");
                return;
            }

            Outlook.Application outlookApp = null;
            try
            {
                outlookApp = new Outlook.Application();
                Outlook.Accounts accounts = outlookApp.Session.Accounts;

                bool isLogged = false;
                foreach (Outlook.Account account in accounts)
                {
                    if (account.SmtpAddress.Equals(txtEmail.Text.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        isLogged = true;
                        break;
                    }
                }

                if (isLogged)
                {
                    List<SqlParameter> parameters = new List<SqlParameter>
                    {
                        new SqlParameter("@Email", txtEmail.Text)
                    };

                    DataTable dt = SQLDataTool.QueryUserData("SELECT * FROM User_Registry WHERE Email=@Email", parameters, PathReader.DVM_link);
                    {
                        if (dt.Rows.Count > 0)
                        {
                            MessageBox.Show("Email exists in Outlook accounts.");
                            MessageBox.Show("User Name: " + dt.Rows[0][4] + "\n" + "Password: " + dt.Rows[0][5]);
                            Close();
                        }
                        else
                        {
                            MessageBox.Show("Please enter the email you are currently logged in.");
                            txtEmail.SelectAll();
                            txtEmail.Focus();
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please enter the email you are currently logged in.");
                    txtEmail.SelectAll();
                    txtEmail.Focus();
                }
            }
            catch (COMException)
            {
                MessageBox.Show("Please open Outlook application before using this feature.");
            }
            finally
            {
                if (outlookApp != null)
                {
                    Marshal.ReleaseComObject(outlookApp);
                }
            }
        }
    }
}