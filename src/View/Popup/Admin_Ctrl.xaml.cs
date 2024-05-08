using System;
using MnS.lib;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Data;
using System.Data.SqlClient;
using System.ComponentModel;

namespace MnS
{
    public partial class Admin_Ctrl : Window
    {
        public Admin_Ctrl()
        {
            InitializeComponent();
            Load_DataSources();
            Closing += AdminCtrl_Closing;
        }

        /// <summary>
        /// 
        /// </summary>
        private void Load_DataSources()
        {
            List<string> dataSources = new List<string>
            {
                "DVM_DT",
                "MNSP_DT",
                "GTData_PM",
                "GTData_CAL"
            };
            comboBox1.ItemsSource = dataSources;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Combo1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboBox1.SelectedItem is string selectedDataSource)
            {
                string dataSource = null;
                switch (selectedDataSource)
                {
                    case "DVM_DT":
                        dataSource = PathReader.DVM_link;
                        break;
                    case "MNSP_DT":
                        dataSource = PathReader.MNSP_link;
                        break;
                    case "GTData_PM":
                        dataSource = PathReader.PM_link;
                        break;
                    case "GTData_CAL":
                        dataSource = PathReader.CAL_link;
                        break;
                    default:
                        break;
                }

                if (dataSource != null)
                {
                    List<string> tables = GetTablesFromDatabase(dataSource);
                    comboBox2.ItemsSource = tables;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataSource"></param>
        /// <returns></returns>
        private List<string> GetTablesFromDatabase(string dataSource)
        {
            List<string> tables = new List<string>();
            try
            {
                DataTable tableNames = SQLDataTool.QueryUserData("SELECT table_name FROM information_schema.tables WHERE table_type = 'BASE TABLE'", new List<SqlParameter>(), dataSource);

                foreach (DataRow row in tableNames.Rows)
                {
                    tables.Add(row["table_name"].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            return tables;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            if (comboBox2.SelectedItem != null)
            {
                string selectedTable = comboBox2.SelectedItem.ToString();
                string selectedDataSource = null;
                string selectedDataSourceName = comboBox1.SelectedItem as string;
                switch (selectedDataSourceName)
                {
                    case "DVM_DT":
                        selectedDataSource = PathReader.DVM_link;
                        break;
                    case "MNSP_DT":
                        selectedDataSource = PathReader.MNSP_link;
                        break;
                    case "GTData_PM":
                        selectedDataSource = PathReader.PM_link;
                        break;
                    case "GTData_CAL":
                        selectedDataSource = PathReader.CAL_link;
                        break;
                    default:
                        MessageBox.Show("Please select a Database.");
                        break;
                }

                if (selectedDataSource != null)
                {
                    try
                    {
                        DataTable tableData = SQLDataTool.QueryUserData($"SELECT * FROM {selectedTable}", new List<SqlParameter>(), selectedDataSource);
                        userGridView.ItemsSource = tableData.DefaultView;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select a Data Table.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            DataTable dataTable = new DataTable();
            string table_name = comboBox2.SelectedItem.ToString();
            foreach (DataGridColumn column in userGridView.Columns)
            {
                if (column is DataGridTextColumn)
                {
                    DataGridTextColumn textColumn = column as DataGridTextColumn;
                    dataTable.Columns.Add(textColumn.Header.ToString());
                }
            }
            foreach (object item in userGridView.Items)
            {
                DataRow newRow = dataTable.NewRow();

                for (int i = 0; i < userGridView.Columns.Count; i++)
                {
                    FrameworkElement cellValue = userGridView.Columns[i].GetCellContent(item);
                    if (cellValue is TextBlock)
                    {
                        newRow[i] = (cellValue as TextBlock).Text;
                    }
                }
                dataTable.Rows.Add(newRow);
            }
            string connectionString = "";

            if (comboBox1.SelectedItem.ToString() == "DVM_DT")
            {
                connectionString = PathReader.DVM_link;
            }
            else if (comboBox1.SelectedItem.ToString() == "MNSP_DT")
            {
                connectionString = PathReader.MNSP_link;
            }
            else if (comboBox1.SelectedItem.ToString() == "GTData_PM")
            {
                connectionString = PathReader.PM_link;
            }
            else if (comboBox1.SelectedItem.ToString() == "GTData_CAL")
            {
                connectionString = PathReader.CAL_link;
            }

            SqlConnection connection = ServerConnection.OpenConnection(connectionString);
            {
                SqlTransaction transaction = connection.BeginTransaction();
                try
                {
                    using (SqlCommand truncateCommand = new SqlCommand($"TRUNCATE TABLE {table_name}", connection, transaction))
                    {
                        truncateCommand.ExecuteNonQuery();
                    }

                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection, SqlBulkCopyOptions.Default, transaction))
                    {
                        bulkCopy.DestinationTableName = table_name;
                        foreach (DataColumn column in dataTable.Columns)
                        {
                            bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                        }

                        bulkCopy.WriteToServer(dataTable);
                    }

                    transaction.Commit();
                    MessageBox.Show("Data saved successfully.");
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    MessageBox.Show($"Error: {ex.Message}");
                }
                finally
                {
                    ServerConnection.CloseConnection(connection);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AdminCtrl_Closing(object sender, CancelEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
        }
    }
}