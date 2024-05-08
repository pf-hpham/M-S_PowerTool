using MnS.lib;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MnS
{
    public partial class Message : Window
    {
        private string email;

        public Message()
        {
            InitializeComponent();
            VersionCheck();
            Assembly assembly = Assembly.GetExecutingAssembly();
            Version version = assembly.GetName().Version;
            vs.Text = version.ToString();
            email = GetLoggedInOutlookEmailAddress();
            til.Text = "THANK YOU FOR USING OUR TOOL!";
            cont.Text = "\tWe know that when you receive this message, our research and development efforts are reaching the people who really need it. That was really a huge encouragement for us.\n\n\tHopefully, the next improvements and upgrades can help your work achieve even more efficiency.\n\n\tCurrently, we are implementing user feedback to upgrade some features and experience better. Please take a moment to complete the survey!\n\n\tBest Regard!";
            note.Text = "Please enter your full name and department and press the Submit button.";
            note_1.Text = "This is an update notification for the new version and only appears in the first time you update or install this tool.";
            note_2.Text = "Currently, we are conducting a survey to gather feedback from user experiences. We aim to enhance and develop this toolkit to be versatile and efficient. Please press the button below to start the survey.";
            name.TextChanged += TextBox_TextChanged;
            des.TextChanged += TextBox_TextChanged;
        }

        private void VersionCheck()
        {
            string filePath = PathReader.config;
            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);
                string ver = "";
                Assembly assembly = Assembly.GetExecutingAssembly();
                Version version = assembly.GetName().Version;

                foreach (string line in lines)
                {
                    string[] part = line.Split(new char[] { '=' }, 2);
                    if (line.StartsWith("Version="))
                    {
                        ver = part[1];
                    }
                }

                if (ver == version.ToString())
                {
                    MiniTool mini = new MiniTool();
                    mini.Show();
                    Close();
                }
            }
            else
            {
                btnClose.IsEnabled = false;
            }
        }

        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            userGridView.Visibility = Visibility.Collapsed;
            SurveyGridView.Visibility = Visibility.Visible;
            btnNext.Visibility = Visibility.Collapsed;
            btnSubmit.Visibility = Visibility.Visible;
        }

        private void btnSubmit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string filePath = @"C:\M+S_Server\User.ini";
                string content = $"username={name.Text}\ndepartment={des.Text}\nemail={email}";

                if (!File.Exists("C:\\M+S_Server\\"))
                {
                    string directoryPath = Path.GetDirectoryName("C:\\M+S_Server\\");
                    if (!Directory.Exists(directoryPath))
                    {
                        Directory.CreateDirectory(directoryPath);
                    }
                    File.Create(filePath).Close();
                }

                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    writer.Write(content);
                }

                MessageBox.Show("Your information has been saved. Now, you can press the Close button to continue using the program\nIf you don't mind, please take a moment to complete the survey to assist us.\n\nPLEASE NOTE: SELECT THE SERVER AGAIN TO UPDATE ALL CHANGES TO YOUR DEVICE AFTER THIS NOTIFICATION!");
                btnClose.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            string source = @"\\pfvn-netapp1\files\87-Maintenance-Services-SEA\_Public\100-M+S_PowerTool\config\Config.ini";
            string filePath = @"C:\M+S_Server\Config.ini";
            File.Copy(source, filePath, true);

            MiniTool mini = new MiniTool();
            mini.Show();

            Close();
        }

        private void CheckTextBoxes()
        {
            if (!string.IsNullOrEmpty(name.Text) && !string.IsNullOrEmpty(des.Text))
            {
                btnSubmit.IsEnabled = true;
            }
            else
            {
                btnSubmit.IsEnabled = false;
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            CheckTextBoxes();
        }

        private void btnServey_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string url = "https://docs.google.com/forms/d/e/1FAIpQLSc_VQtZljwIgXmWnLEfrox5847sHOTXZyf_7J7fw64tRWLtEQ/viewform";
                Process.Start(url);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        public string GetLoggedInOutlookEmailAddress()
        {
            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
            Outlook.Accounts accounts = outlookNamespace.Accounts;

            foreach (Outlook.Account account in accounts)
            {
                if (account.SmtpAddress != null && account.SmtpAddress != "")
                {
                    return account.SmtpAddress;
                }
            }
            return null;
        }

        #region Focus
        private void name_GotFocus(object sender, RoutedEventArgs e)
        {
            if (name.Text == "Your fullname")
            {
                name.Text = "";
            }
        }

        private void name_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(name.Text))
            {
                name.Text = "Your fullname";
            }
        }

        private void des_GotFocus(object sender, RoutedEventArgs e)
        {
            if (des.Text == "Your department")
            {
                des.Text = "";
            }
        }

        private void des_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(des.Text))
            {
                des.Text = "Your department";
            }
        }
        #endregion
    }
}