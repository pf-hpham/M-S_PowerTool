using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using MnS.lib;

namespace MnS
{
    public partial class MiniTool : Window
    {
        #region Variable
        public static bool loggedin = false;
        public static string server;
        #endregion

        public MiniTool()
        {
            InitializeComponent();
            status.Fill = Brushes.Red;
            server = "";
            ODBC_Check();
            Create_Server();
            Server_Checked();
            PathReader.Read_Server();
            Top = SystemParameters.WorkArea.Height - Height;
            Width = SystemParameters.PrimaryScreenWidth;
            mnApp.SelectedIndex = 0;
        }

        private void App_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void MnExit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void MnMnS_Click(object sender, RoutedEventArgs e)
        {
            if (loggedin)
            {
                MainWindow mainWindow = new MainWindow();
                mainWindow.Show();
            }
            else
            {
                Registry register = new Registry();
                register.Show();
            }
        }

        private void MnMvex_Click(object sender, RoutedEventArgs e)
        {
            Movex routing = new Movex();
            routing.Show();
        }

        private void MnEDM_Click(object sender, RoutedEventArgs e)
        {
            EDM edm = new EDM();
            edm.Show();
        }

        private void MnMB_Click(object sender, RoutedEventArgs e)
        {
            DateConvert datepicker = new DateConvert();
            datepicker.Show();
        }

        private void MnMess_Click(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(PathReader.boxplot_foldermes))
            {
                Directory.CreateDirectory(PathReader.boxplot_foldermes);
            }

            MessData messdata = new MessData();
            messdata.Show();
        }

        private void MnApp_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBoxItem selectedApp = (ComboBoxItem)mnApp.SelectedItem;
            if (selectedApp != null)
            {
                string selectedAppName = selectedApp.Content.ToString();
                switch (selectedAppName)
                {
                    case "EDM Documents":
                        mnContent.ItemsSource = new List<string> { "All file", "Pdf file", "Doc file", "Excel file" };
                        mnContent.SelectedIndex = 0;
                        break;
                    case "Web Browser":
                        mnContent.ItemsSource = new List<string> { "Google Search", "Misumi", "RS Component" , "ezPortal" , "Techmaster Portal" };
                        mnContent.SelectedIndex = 0;
                        break;
                    default:
                        break;
                }
            }
        }

        private void MnInput_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                SearchButton_Click(sender, e);
            }
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedApp = mnApp.Text;
            string selectedContent = mnContent.Text;
            string searchText = mnInput.Text.Trim();
            string searchUrl = "";
            switch (selectedApp)
            {
                case "Web Browser":
                    switch (selectedContent)
                    {
                        case "Google Search":
                            searchUrl = PathReader.GGSearch + searchText.Replace(" ", "+");
                            break;
                        case "RS Component":
                            searchUrl = PathReader.RSComponent + searchText.Replace(" ", "+");
                            break;
                        case "Misumi":
                            searchUrl = PathReader.Misumi + searchText.Replace(" ", "-");
                            break;
                        case "ezPortal":
                            searchUrl = PathReader.EzPortal;
                            break;
                        case "Techmaster Portal":
                            searchUrl = PathReader.Techmaster;
                            break;
                        default:
                            break;
                    }
                    Process.Start("msedge.exe", searchUrl);
                    break;
                case "EDM Documents":
                    if (selectedContent == "Pdf file")
                    {
                        Process.Start(PathReader.EDM_link + searchText + "\\pdf");
                    }
                    else
                    {
                        Process.Start(PathReader.EDM_link + searchText);
                    }
                    break;
                default:
                    break;
            }
        }

        private void MnClock_Click(object sender, RoutedEventArgs e)
        {
            Clock clock = new Clock();
            clock.Show();
        }

        private void Guide_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(PathReader.Guide_link);
        }

        private void MnResize_Click(object sender, RoutedEventArgs e)
        {
            Icon icon = new Icon();
            icon.Show();
            Hide();
        }

        #region Server
        private void Create_Server()
        {
            if (!File.Exists(PathReader.server))
            {
                string directoryPath = Path.GetDirectoryName(PathReader.server);
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }
                File.Create(PathReader.server).Close();
                File.WriteAllText(PathReader.server, "Server=None");
            }
        }

        private void ServerChange_Click(object sender, RoutedEventArgs e)
        {
            ServerChange svChange = new ServerChange();
            svChange.Closed += Server_Check;
            svChange.Show();
        }

        public void Server_Check(object sender, EventArgs e)
        {
            Server_Checked();
        }

        public void Server_Checked()
        {
            if (PathReader.server != "")
            {
                string[] lines = File.ReadAllLines(PathReader.server);
                foreach (string line in lines)
                {
                    if (line.Contains("Server="))
                    {
                        string[] part = line.Split(new char[] { '=' }, 2);
                        if (part[1].ToString() == "None")
                        {
                            status.Fill = Brushes.Red;
                            MessageBox.Show("Server has not been initialized.\nPlease select Server to update the software.\nClick to the National Flag or Country Name to choose Server.");
                            ServerChange server = new ServerChange();
                            server.Closed += Server_Check;
                            server.Show();
                            break;
                        }
                        if (part[1].ToString() == "VIE")
                        {
                            region_sv.Content = "VIE";
                            server = "VIE";
                            Uri newImageUri = new Uri("/assets/images/VIE.png", UriKind.Relative);
                            mnSea.Source = new BitmapImage(newImageUri);
                            break;
                        }
                        else if (part[1].ToString() == "SGP")
                        {
                            region_sv.Content = "SGP";
                            server = "SGP";
                            Uri newImageUri = new Uri("/assets/images/SGP.png", UriKind.Relative);
                            mnSea.Source = new BitmapImage(newImageUri);
                            break;
                        }
                        if (part[1].ToString() == "IDN")
                        {
                            region_sv.Content = "IDN";
                            server = "IDN";
                            Uri newImageUri = new Uri("/assets/images/IDN.png", UriKind.Relative);
                            mnSea.Source = new BitmapImage(newImageUri);
                            break;
                        }
                    }
                }
            }
            else
            {
                status.Fill = Brushes.Red;
                MessageBox.Show("Error: Server is not exits or connected, please contact administator.");
            }
        }
        #endregion

        #region ODBC Check
        private void ODBC_Check()
        {
            string connectionString = "DSN=odssg_vn";
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    connection.Open();
                    Console.WriteLine("ODBC connection to the server is successful.");
                    status.Fill = Brushes.LightGreen;
                    connection.Close();
                }
            }
            catch (OdbcException ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                status.Fill = Brushes.Red;
                MessageBoxResult result = MessageBox.Show("Could not establish ODBC connection to the server.\nDo you want to proceed with creating the ODBC connection?", "Connection Error", MessageBoxButton.YesNo);

                if (result == MessageBoxResult.Yes)
                {
                    string powerShellScriptPath = @"\\pfvn-netapp1\files\87-Maintenance-Services-SEA\_Public\100-M+S_PowerTool\odssg_vn.ps1";
                    RunPowerShellScript(powerShellScriptPath);
                }
                else
                {
                    MessageBox.Show("Skip this connection.");
                }
            }
        }

        private void RunPowerShellScript(string scriptPath)
        {
            try
            {
                ProcessStartInfo processInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-ExecutionPolicy Bypass -File \"{scriptPath}\"",
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    CreateNoWindow = false,
                    UseShellExecute = false
                };

                using (Process process = new Process())
                {
                    process.StartInfo = processInfo;
                    process.Start();
                    process.WaitForExit();

                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    Console.WriteLine($"Output: {output}");
                    Console.WriteLine($"Error: {error}");
                }
                MessageBox.Show("Created the ODBC connection successfull.");
                ODBC_Check();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error running PowerShell script: {ex.Message}");
            }
        }
        #endregion

        private void Feedback_Click(object sender, RoutedEventArgs e)
        {
            OutlookTool.Feedback("hpham", "dhoang");
        }

        private void MnPixel_Click(object sender, RoutedEventArgs e)
        {
            Pixel dectect = new Pixel();
            dectect.Show();
        }

        private void MnText_Click(object sender, RoutedEventArgs e)
        {
            AutoRecord text = new AutoRecord();
            text.Show();
        }

        private void MnSQL_Click(object sender, RoutedEventArgs e)
        {
            SQL_cmd sql = new SQL_cmd();
            sql.Show();
        }
    }
}