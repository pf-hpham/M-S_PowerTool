using System;
using System.IO;
using MnS.lib;
using System.Windows;
using System.Windows.Media;

namespace MnS
{
    public partial class ServerChange : Window
    {
        public static string SEA_regional;

        public ServerChange()
        {
            UserLogTool.UserData("Change server");
            InitializeComponent();
            Server_Check();
            LoadAndDisplayConfigData();
        }

        private void LoadAndDisplayConfigData()
        {
            try
            {
                string filePath = PathReader.config;
                if (File.Exists(filePath))
                {
                    string[] lines = File.ReadAllLines(filePath);
                    string vnContent = "";
                    string sgContent = "";
                    string idContent = "";
                    foreach (string line in lines)
                    {
                        string[] part = line.Split(new char[] { '=' }, 2);
                        if (line.StartsWith("VIE="))
                        {
                            vnContent += $"• {part[1].Replace("\"", "")}" + Environment.NewLine;
                        }
                        if (line.StartsWith("SGP="))
                        {
                            sgContent += $"• {part[1].Replace("\"", "")}" + Environment.NewLine;
                        }
                        if (line.StartsWith("IDN="))
                        {
                            idContent += $"• {part[1].Replace("\"", "")}" + Environment.NewLine;
                        }
                    }
                    VN_function.Text = vnContent;
                    SG_function.Text = sgContent;
                    ID_function.Text = idContent;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex + "\nClick OK to create new Config file.");
            }
        }

        public void Server_Check()
        {
            if (PathReader.server != "")
            {
                string[] lines = File.ReadAllLines(PathReader.server);
                foreach (string line in lines)
                {
                    if (line.Contains("Server="))
                    {
                        string[] part = line.Split(new char[] { '=' }, 2);
                        if (part[1].ToString() == "VIE")
                        {
                            MessageBox.Show("Server in system is VietNam.");
                            vn_border.Background = new SolidColorBrush(Colors.DarkCyan);
                            break;
                        }
                        else if (part[1].ToString() == "SGP")
                        {
                            MessageBox.Show("Server in system is Singapore.");
                            sg_border.Background = new SolidColorBrush(Colors.DarkCyan);
                            break;
                        }
                        if (part[1].ToString() == "IDN")
                        {
                            MessageBox.Show("Server in system is Indonesia.");
                            id_border.Background = new SolidColorBrush(Colors.DarkCyan);
                            break;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Server has not been initialized.\nPlease select Server to update the software.\nClick to the National Flag or Country Name to choose Server.");
            }
        }

        private void SGP_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                string filepath = "C:\\M+S_Server\\Path_Server.ini";
                string movex = "C:\\M+S_Server\\Movex_Server.ini";
                string filesource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Path_SGP.ini";
                string movexsource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Movex_SGP.ini";
                File.Copy(filesource, filepath, true);
                File.Copy(movexsource, movex, true);
                MessageBox.Show($"Created new:\n{filepath}\n{movex}\nsuccesfully.");

                string[] lines = File.ReadAllLines(PathReader.server);
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].StartsWith("Server="))
                    {
                        lines[i] = "Server=SGP";
                        sg_border.Background = new SolidColorBrush(Colors.DarkCyan);
                        id_border.Background = new SolidColorBrush(Colors.Transparent);
                        vn_border.Background = new SolidColorBrush(Colors.Transparent);
                        break;
                    }
                }
                File.WriteAllLines(PathReader.server, lines);
                MessageBox.Show("Server changed to SGP-MF1\nPlease restart this program to update all server configuration.");
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating server: {ex.Message}");
            }
        }

        private void VIE_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                string filepath = "C:\\M+S_Server\\Path_Server.ini";
                string movex = "C:\\M+S_Server\\Movex_Server.ini";
                string filesource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Path_VIE.ini";
                string movexsource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Movex_VIE.ini";
                File.Copy(filesource, filepath, true);
                File.Copy(movexsource, movex, true);
                MessageBox.Show($"Created new:\n{filepath}\n{movex}\nsuccesfully.");

                string[] lines = File.ReadAllLines(PathReader.server);
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].StartsWith("Server="))
                    {
                        lines[i] = "Server=VIE";
                        sg_border.Background = new SolidColorBrush(Colors.Transparent);
                        id_border.Background = new SolidColorBrush(Colors.Transparent);
                        vn_border.Background = new SolidColorBrush(Colors.DarkCyan);
                        break;
                    }
                }
                File.WriteAllLines(PathReader.server, lines);
                MessageBox.Show("Server updated to VIE=VN1\nPlease restart this program to update all server configuration.");
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating server: {ex.Message}");
            }
        }

        private void IDN_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                string filepath = "C:\\M+S_Server\\Path_Server.ini";
                string movex = "C:\\M+S_Server\\Movex_Server.ini";
                string filesource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Path_IDN.ini";
                string movexsource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Movex_IDN.ini";
                File.Copy(filesource, filepath, true);
                File.Copy(movexsource, movex, true);
                MessageBox.Show($"Created new:\n{filepath}\n{movex}\nsuccesfully.");

                string[] lines = File.ReadAllLines(PathReader.server);
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].StartsWith("Server="))
                    {
                        lines[i] = "Server=IDN";
                        sg_border.Background = new SolidColorBrush(Colors.Transparent);
                        id_border.Background = new SolidColorBrush(Colors.DarkCyan);
                        vn_border.Background = new SolidColorBrush(Colors.Transparent);
                        break;
                    }
                }
                File.WriteAllLines(PathReader.server, lines);
                MessageBox.Show("Server updated to IDN-BT1\nPlease restart this program to update all server configuration.");
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating server: {ex.Message}");
            }
        }

        private void ConfigChange_Click(object sender, RoutedEventArgs e)
        {
            Config_property configure = new Config_property(PathReader.config);
            configure.Show();
        }

        private void MovexChange_Click(object sender, RoutedEventArgs e)
        {
            Config_property movex = new Config_property(PathReader.Movex);
            movex.Show();
        }

        private void AppChange_Click(object sender, RoutedEventArgs e)
        {
            Config_property app = new Config_property(PathReader.filePath);
            app.Show();
        }

        private void ImageSG_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            SGP_Click(sender, e);
        }

        private void ImagVN_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            VIE_Click(sender, e);
        }

        private void ImagID_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            IDN_Click(sender, e);
        }
    }
}