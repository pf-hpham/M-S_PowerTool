using Microsoft.Win32;
using MnS.lib;
using System;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace MnS
{
    public partial class SQL_cmd : Window
    {
        #region Variable
        private bool flag;
        private string cmd;
        private OdbcConnection cnt;
        private SqlConnection sql_cnt;

        private DataTable temp_table;
        #endregion

        public SQL_cmd()
        {
            UserLogTool.UserData("Using SQL function");
            InitializeComponent();

            flag = false;
        }

        #region Button
        private void ConnectDB_Clicked(object sender, RoutedEventArgs e)
        {
            if (!flag)
            {
                if (sql_db.SelectedIndex == 0)
                {
                    cnt = new OdbcConnection
                    {
                        ConnectionString = PathReader.ODSSG_server
                    };

                    try
                    {
                        cnt.Open();
                        sql_connect.Content = "CONNECTED";
                        sql_connect.Background = Brushes.Green;
                        flag = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Failed to connect: {ex.Message}");
                    }
                }
                else if (sql_db.SelectedIndex == 1)
                {
                    cnt = new OdbcConnection
                    {
                        ConnectionString = PathReader.EDM_server
                    };

                    try
                    {
                        cnt.Open();
                        sql_connect.Content = "CONNECTED";
                        sql_connect.Background = Brushes.Green;
                        flag = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Failed to connect: {ex.Message}");
                    }
                }
                else if (sql_db.SelectedIndex == 2)
                {
                    if (sql_dbus.Text == "GTLogin" && sql_dbpw.Password == "PASSword1")
                    {
                        sql_cnt = new SqlConnection
                        {
                            ConnectionString = PathReader.PM_link
                        };

                        try
                        {
                            sql_cnt.Open();
                            sql_connect.Content = "CONNECTED";
                            sql_connect.Background = Brushes.Green;
                            flag = true;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Failed to connect: {ex.Message}");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Wrong User or Password, please try again!");
                    }
                }
                else if (sql_db.SelectedIndex == 3)
                {
                    if (sql_dbus.Text == "GTLogin" && sql_dbpw.Password == "PASSword1")
                    {
                        sql_cnt = new SqlConnection
                        {
                            ConnectionString = PathReader.CAL_link
                        };

                        try
                        {
                            sql_cnt.Open();
                            sql_connect.Content = "CONNECTED";
                            sql_connect.Background = Brushes.Green;
                            flag = true;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Failed to connect: {ex.Message}");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Wrong User or Password, please try again!");
                    }
                }
            }
            else
            {
                try
                {
                    if (cnt != null)
                    {
                        cnt.Close();
                    }
                    else if (sql_cnt != null)
                    {
                        sql_cnt.Close();
                    }
                    sql_connect.Content = "DISCONNECTED";
                    sql_connect.Background = Brushes.Red;
                    flag = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to disconnect: {ex.Message}");
                }
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            ExcelTool.ExportExcelWithDialog(temp_table, "SQL Data", "SQL Data");
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (flag)
                {
                    if (sql_db.SelectedIndex == 0 || sql_db.SelectedIndex == 1)
                    {
                        OdbcCommand cmd = new OdbcCommand
                        {
                            Connection = cnt,
                            CommandType = CommandType.Text,
                            CommandText = sql_cmd.Text
                        };

                        OdbcDataAdapter data = new OdbcDataAdapter
                        {
                            SelectCommand = cmd
                        };
                        temp_table = new DataTable();
                        data.Fill(temp_table);

                        DateTimeFormat.DatetimeFormat(temp_table, sql_grid, "A");
                    }
                    else if (sql_db.SelectedIndex == 2 || sql_db.SelectedIndex == 3)
                    {
                        SqlCommand cmd = new SqlCommand
                        {
                            Connection = sql_cnt,
                            CommandType = CommandType.Text,
                            CommandText = sql_cmd.Text
                        };

                        SqlDataAdapter data = new SqlDataAdapter
                        {
                            SelectCommand = cmd
                        };
                        temp_table = new DataTable();
                        data.Fill(temp_table);

                        DateTimeFormat.DatetimeFormat(temp_table, sql_grid, "A");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Load_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == true)
                {
                    string filePath = openFileDialog.FileName;
                    string[] lines = File.ReadAllLines(filePath);

                    foreach (string line in lines)
                    {
                        if (line.Contains("cmd="))
                        {
                            string[] parts = line.Split(new char[] { '=' }, 2);
                            if (parts.Length == 2)
                            {
                                cmd = parts[1];
                            }
                        }
                    }
                }

                if (cmd != "")
                {
                    sql_cmd.Text = cmd;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "Text files (*.txt)|*.txt";
            saveFileDialog.FileName = "sql_commander.txt";

            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;

                try
                {
                    using (StreamWriter writer = new StreamWriter(filePath))
                    {
                        writer.WriteLine("cmd=" + sql_cmd.Text);
                    }

                    MessageBox.Show("File saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        #endregion

        private void Server_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sql_db.SelectedIndex == 0 || sql_db.SelectedIndex == 1) && sql_dbus != null && sql_dbpw != null)
            {
                sql_dbus.IsEnabled = false;
                sql_dbpw.IsEnabled = false;
                sql_dbdriver.SelectedIndex = 0;
            }
            else if (sql_db.SelectedIndex == 2 || sql_db.SelectedIndex == 3)
            {
                sql_dbus.IsEnabled = true;
                sql_dbpw.IsEnabled = true;
                sql_dbdriver.SelectedIndex = 1;
            }
        }

        private void CMD_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sql_cmdref.SelectedIndex == 1)
            {
                string filePath = @"\\pfvn-netapp1\files\87-Maintenance-Services-SEA\_Public\100-M+S_PowerTool\sql_commander\sql_routing.txt";
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.Contains("cmd="))
                        {
                            string[] parts = line.Split(new char[] { '=' }, 2);
                            if (parts.Length == 2)
                            {
                                sql_cmd.Text = parts[1];
                            }
                        }
                    }
                }
            }
            else if (sql_cmdref.SelectedIndex == 2)
            {
                string filePath = @"\\pfvn-netapp1\files\87-Maintenance-Services-SEA\_Public\100-M+S_PowerTool\sql_commander\sql_bom.txt";
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.Contains("cmd="))
                        {
                            string[] parts = line.Split(new char[] { '=' }, 2);
                            if (parts.Length == 2)
                            {
                                sql_cmd.Text = parts[1];
                            }
                        }
                    }
                }
            }
        }
    }
}