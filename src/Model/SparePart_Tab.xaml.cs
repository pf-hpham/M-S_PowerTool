using System;
using System.IO;
using MnS.lib;
using System.Data;
using System.Linq;
using LiveCharts;
using Microsoft.Win32;
using LiveCharts.Wpf;
using System.Windows;
using System.Diagnostics;
using System.Windows.Media;
using System.Windows.Input;
using System.Data.SqlClient;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Windows.Media.Imaging;

namespace MnS
{
    public partial class SparePart_Tab : UserControl
    {
        #region
        private DataTable spImportTable;
        private DataTable SPtempHistory;
        private DataTable SPtopup;
        private readonly DataTable SPIMRTable;
        private readonly DataTable SPEXPTable;
        #endregion

        public SparePart_Tab()
        {
            UserLogTool.UserData("Using Spare Part function");
            InitializeComponent();
            SPEXPTable = SPexpTable();
            SPIMRTable = SPimrTable();
            LoadTPU();
            LoadItemID();
        }

        private DataTable SPDatabase()
        {
            DataTable SPTable = SQLDataTool.QueryUserData("SELECT * FROM Device_List", new List<SqlParameter>(), PathReader.MNSP_link);
            return SPTable;
        }

        private async void SPload_Click(object sender, RoutedEventArgs e)
        {
            loadingProg.Visibility = Visibility.Visible;
            loadingProg.Value = 0;

            await ProgressBarTool.ProgressBarAsync(SPDatabase(), (progress) =>
            {
                ProgressBarTool.UpdateProgressBar(loadingProg, progress);
            });
            loadingProg.Visibility = Visibility.Hidden;

            DateTimeFormat.DatetimeFormat(SPDatabase(), spDataGridView, "A");

            DataColumnCollection columns = SPDatabase().Columns;
            if (columns is null)
            {
                throw new ArgumentNullException(nameof(columns));
            }
            foreach (DataColumn column in SPDatabase().Columns)
            {
                spComboBox.Items.Add(column.ColumnName);
            }

            spInforGridView.Visibility = Visibility.Visible;
            MessageBox.Show("Loaded Database successfully!");
        }

        private void SpTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string searchText = spTextBox.Text;

            DataView dataView = SPDatabase().DefaultView;

            if (searchText != null)
            {
                string selectedColumn = spComboBox.SelectedValue as string;

                if (!string.IsNullOrWhiteSpace(selectedColumn))
                {
                    DataColumn column = dataView.Table.Columns[selectedColumn];

                    dataView.RowFilter = column.DataType == typeof(DateTime)
                        ? DateTime.TryParse(searchText, out DateTime searchDate) ? $"{selectedColumn} = #{searchDate.ToShortDateString()}#" : "1=0"
                        : column.DataType == typeof(int)
                            ? int.TryParse(searchText, out int searchNumber) ? $"{selectedColumn} = {searchNumber}" : "1=0"
                            : $"{selectedColumn} LIKE '%{searchText}%'";
                }
            }
            spDataGridView.ItemsSource = dataView;
        }

        private void SpExportExcel(object sender, RoutedEventArgs e)
        {
            ExcelTool.ExportExcelWithDialog(SPDatabase(), "SP_Data", "Spare_ExportData");
        }

        private void SpSendDatabasetoOutlook(object sender, RoutedEventArgs e)
        {
            OutlookTool.SendEmailWithExcelAttachment(SPDatabase());
        }

        private void SpDataGridView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (spDataGridView.SelectedValue is DataRowView selectedRowView)
            {
                DataRow selectedRow = selectedRowView.Row;
                string idItem = selectedRow["ItemID"].ToString();
                if (selectedRow != null)
                {
                    string modelNo = selectedRow["Model"].ToString();
                    string description = selectedRow["Description"].ToString();
                    string department = selectedRow["TPU"].ToString();
                    string storeLoc = selectedRow["Store_loc"].ToString();
                    string quantity = selectedRow["Quantity"].ToString();
                    string safetyStock = selectedRow["Safety_stock"].ToString();

                    spID.Text = idItem;
                    spModel.Text = modelNo;
                    spDescription.Text = description;
                    spDepart.Text = department;
                    spStore.Text = storeLoc;
                    spQuantity.Text = quantity;
                    spSafety.Text = safetyStock;
                }
                LoadChartData(idItem);

                DataTable SPImage_Table = SQLDataTool.QueryUserData($"SELECT * FROM Image_List WHERE ItemID = '{selectedRow["ItemID"]}'", new List<SqlParameter>(), PathReader.MNSP_link);
                DataTable SPQuote_Table = SQLDataTool.QueryUserData($"SELECT * FROM Quote_List WHERE ItemID = '{selectedRow["ItemID"]}'", new List<SqlParameter>(), PathReader.MNSP_link);

                if (SPImage_Table.Rows.Count > 0)
                {
                    string imagePath = Path.Combine(PathReader.Image_device, SPImage_Table.Rows[0]["Image_name"].ToString());
                    if (File.Exists(imagePath))
                    {
                        BitmapImage bitmapImage = new BitmapImage(new Uri(imagePath));
                        spDevice_image.Source = bitmapImage;
                    }
                    else
                    {
                        spDevice_image.Source = null;
                    }
                }
                else
                {
                    spDevice_image.Source = null;
                }
                sphyperlinkText.Text = SPQuote_Table.Rows.Count > 0 ? SPQuote_Table.Rows[0]["Quotation_name"].ToString() : null;
            }
        }

        private void OpenFileHyperlink_Click(object sender, RoutedEventArgs e)
        {
            if (spDataGridView.SelectedValue is DataRowView selectedRowView)
            {
                DataRow selectedRow = selectedRowView.Row;
                string itemId = selectedRow["ItemID"].ToString();

                DataTable SPQuote_Table = SQLDataTool.QueryUserData($"SELECT Quotation_name FROM Quote_List WHERE ItemID = '{itemId}'", new List<SqlParameter>(), PathReader.MNSP_link);

                if (SPQuote_Table.Rows.Count > 0)
                {
                    string fileName = SPQuote_Table.Rows[0]["Quotation_name"].ToString();

                    if (!string.IsNullOrEmpty(fileName))
                    {
                        string filePath = Path.Combine(PathReader.Quote_folder, fileName);
                        if (File.Exists(filePath))
                        {
                            try
                            {
                                Process.Start(filePath);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error opening file: " + ex.Message);
                            }
                        }
                        else
                        {
                            MessageBox.Show("File not found.");
                        }
                    }
                    else
                    {
                        MessageBox.Show("File name is missing.");
                    }
                }
                else
                {
                    MessageBox.Show("No data available for this item.");
                }
            }
        }

        private void LoadChartData(string selectedRow)
        {
            DataTable exportData = SQLDataTool.QueryUserData("SELECT ItemID, Export_quantity, Export_date FROM Export_History WHERE ItemID = @selectedRow", new List<SqlParameter> { new SqlParameter("@selectedRow", selectedRow) }, PathReader.MNSP_link);
            DataTable importData = SQLDataTool.QueryUserData("SELECT ItemID, Import_quantity, Import_date FROM Import_History WHERE ItemID = @selectedRow", new List<SqlParameter> { new SqlParameter("@selectedRow", selectedRow) }, PathReader.MNSP_link);

            List<KeyValuePair<string, int>> chartData = new List<KeyValuePair<string, int>>();
            foreach (DataRow row in importData.Rows)
            {
                DateTime date = (DateTime)row["Import_date"];
                int quantity = (int)row["Import_quantity"];
                string value = "Import\n" + date.ToShortDateString();
                chartData.Add(new KeyValuePair<string, int>(value, quantity));
            }
            foreach (DataRow row in exportData.Rows)
            {
                DateTime date = (DateTime)row["Export_date"];
                int quantity = -(int)row["Export_quantity"];
                string value = "Export\n" + date.ToShortDateString();
                chartData.Add(new KeyValuePair<string, int>(value, quantity));
            }

            chartData.Sort((pair1, pair2) => DateTime.Parse(pair1.Key.Split('\n')[1]).CompareTo(DateTime.Parse(pair2.Key.Split('\n')[1])));
            ColumnSeries importExportSeries = new ColumnSeries
            {
                Title = "Item: " + selectedRow,
                FontSize = 13,
                Values = new ChartValues<int>(chartData.Select(dp => dp.Value)),
                DataLabels = true,
                Fill = Brushes.OrangeRed,
            };

            spChart.AxisX.Clear();
            spChart.AxisX.Add(new Axis
            {
                Title = "",
                Labels = new List<string>(chartData.Select(dp => dp.Key)),
                LabelsRotation = 0,
                Foreground = Brushes.Black
            });

            spChart.AxisY.Clear();
            spChart.AxisY.Add(new Axis
            {
                Title = "Item: " + selectedRow,
                FontSize = 13,
                Foreground = Brushes.Black
            });

            spChart.Series.Clear();
            spChart.Series.Add(importExportSeries);
        }

        private void LoadTPU()
        {
            try
            {
                ComboBoxTool.LoadCombobox("SELECT DISTINCT TPU FROM Location", new List<SqlParameter>(), PathReader.MNSP_link, spimptpu, "TPU", "TPU", -1);
                ComboBoxTool.LoadCombobox("SELECT DISTINCT Cabinet FROM Store_loc", new List<SqlParameter>(), PathReader.MNSP_link, spstrloc, "Cabinet", "Cabinet", -1);
                ComboBoxTool.LoadCombobox("SELECT * FROM Unit", new List<SqlParameter>(), PathReader.MNSP_link, spimpunit, "Unit", "Unit", -1);
                ComboBoxTool.LoadCombobox("SELECT * FROM Currency", new List<SqlParameter>(), PathReader.MNSP_link, spimpcurrency, "Currency_code", "Currency_code", -1);

                DataTable userTable = SQLDataTool.QueryUserData($"SELECT FullName FROM User_Registry WHERE UserID = '{UserLogTool.userInfor.Rows[0][0]}'", new List<SqlParameter>(), PathReader.DVM_link);
                spimpt.Text = userTable.Rows[0]["FullName"].ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while loading departments: " + ex.Message);
            }
        }

        private void Spimptpu_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (spimptpu.SelectedValue != null)
            {
                string selectedTPU = spimptpu.SelectedValue.ToString();
                try
                {
                    ComboBoxTool.LoadCombobox($"SELECT * FROM Location WHERE TPU = '{selectedTPU}'", new List<SqlParameter>(), PathReader.MNSP_link, spimploc, "Line", "Line", -1);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error while loading locations: " + ex.Message);
                }
            }
        }

        private void Spstrloc_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (spstrloc.SelectedValue != null)
            {
                string selectedLoc = spstrloc.SelectedValue.ToString();
                try
                {
                    ComboBoxTool.LoadCombobox($"SELECT * FROM Store_loc WHERE Cabinet = '{selectedLoc}'", new List<SqlParameter>(), PathReader.MNSP_link, spstrlocno, "Loc_no", "Loc_no", -1);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error while loading locations: " + ex.Message);
                }
            }
        }

        private DataTable SPnewdevicetable()
        {
            DataTable SPimporttable = new DataTable();
            SPimporttable.Columns.Add("ItemID", typeof(string));
            SPimporttable.Columns.Add("Description", typeof(string));
            SPimporttable.Columns.Add("Model", typeof(string));
            SPimporttable.Columns.Add("Part_no", typeof(string));
            SPimporttable.Columns.Add("Manufacturer", typeof(string));
            SPimporttable.Columns.Add("TPU", typeof(string));
            SPimporttable.Columns.Add("Location", typeof(string));
            SPimporttable.Columns.Add("Machine", typeof(string));
            SPimporttable.Columns.Add("Store_loc", typeof(string));
            SPimporttable.Columns.Add("Loc_no", typeof(string));
            SPimporttable.Columns.Add("Quantity", typeof(int));
            SPimporttable.Columns.Add("Safety_stock", typeof(int));
            SPimporttable.Columns.Add("Unit", typeof(string));
            SPimporttable.Columns.Add("Price", typeof(decimal));
            SPimporttable.Columns.Add("Currency", typeof(string));
            SPimporttable.Columns.Add("Register_date", typeof(DateTime));
            SPimporttable.Columns.Add("Register", typeof(string));
            SPimporttable.Columns.Add("Noted", typeof(string));
            SPimporttable.Columns.Add("Image_name", typeof(string));
            SPimporttable.Columns.Add("Quotation_name", typeof(string));

            spImport.ItemsSource = SPimporttable.DefaultView;
            return SPimporttable;
        }

        private void LoadItemID()
        {
            try
            {
                ComboBoxTool.LoadCombobox("SELECT ItemID FROM Device_List", new List<SqlParameter>(), PathReader.MNSP_link, spimrid, "ItemID", "ItemID", -1);
                ComboBoxTool.LoadCombobox("SELECT ItemID FROM Device_List", new List<SqlParameter>(), PathReader.MNSP_link, spexpid, "ItemID", "ItemID", -1);

                DataTable userTable = SQLDataTool.QueryUserData($"SELECT FullName FROM User_Registry WHERE UserID = '{UserLogTool.userInfor.Rows[0][0]}'", new List<SqlParameter>(), PathReader.DVM_link);
                spimrimt.Text = userTable.Rows[0]["FullName"].ToString();
                spexpimt.Text = userTable.Rows[0]["FullName"].ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while loading departments: " + ex.Message);
            }
        }

        private void SpbtnImport_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(spimpid.Text) ||
                string.IsNullOrWhiteSpace(spimpdes.Text) ||
                string.IsNullOrWhiteSpace(spimpmd.Text) ||
                string.IsNullOrWhiteSpace(spimpmanu.Text) ||
                string.IsNullOrWhiteSpace(spimptpu.Text) ||
                string.IsNullOrWhiteSpace(spimploc.Text) ||
                string.IsNullOrWhiteSpace(spimpmc.Text) ||
                string.IsNullOrWhiteSpace(spstrloc.Text) ||
                string.IsNullOrWhiteSpace(spstrlocno.Text) ||
                string.IsNullOrWhiteSpace(spimppn.Text) ||
                string.IsNullOrWhiteSpace(spimpquan.Text) ||
                string.IsNullOrWhiteSpace(spimpunit.Text) ||
                string.IsNullOrWhiteSpace(spimpcurrency.Text) ||
                string.IsNullOrWhiteSpace(spimpsaty.Text) ||
                string.IsNullOrWhiteSpace(spimpprice.Text) ||
                string.IsNullOrWhiteSpace(spdate.Text) ||
                string.IsNullOrWhiteSpace(spimpt.Text) ||
                string.IsNullOrWhiteSpace(spnewimg.Text) ||
                string.IsNullOrWhiteSpace(sptmpnote.Text))
            {
                MessageBox.Show("Please fill in all required fields or enter 'tbd' or 'null' if unknown.");
            }
            else
            {
                if (spImportTable == null)
                {
                    spImportTable = SPnewdevicetable();
                }

                bool isDuplicate = false;
                foreach (DataRow row in spImportTable.Rows)
                {
                    if (row["ItemID"].ToString() == spimpid.Text.Trim())
                    {
                        isDuplicate = true;
                        break;
                    }
                }

                foreach (DataRow row in SPDatabase().Rows)
                {
                    if (row["ItemID"].ToString() == spimpid.Text.Trim())
                    {
                        isDuplicate = true;
                        break;
                    }
                }

                if (!isDuplicate)
                {
                    DataRow newRow = spImportTable.NewRow();
                    newRow["ItemID"] = spimpid.Text;
                    newRow["Description"] = spimpdes.Text;
                    newRow["Model"] = spimpmd.Text;
                    newRow["Part_no"] = spimppn.Text;
                    newRow["Manufacturer"] = spimpmanu.Text;
                    newRow["TPU"] = spimptpu.Text;
                    newRow["Location"] = spimploc.Text;
                    newRow["Machine"] = spimpmc.Text;
                    newRow["Store_loc"] = spstrloc.Text;
                    newRow["Loc_no"] = spstrlocno.Text;
                    newRow["Quantity"] = spimpquan.Text;
                    newRow["Safety_stock"] = spimpsaty.Text;
                    newRow["Unit"] = spimpunit.Text;
                    newRow["Price"] = spimpprice.Text;
                    newRow["Currency"] = spimpcurrency.Text;
                    newRow["Register_date"] = spdate.SelectedDate;
                    newRow["Register"] = spimpt.Text;
                    newRow["Noted"] = sptmpnote.Text;
                    newRow["Image_name"] = spnewimg.Text;
                    newRow["Quotation_name"] = spnewquote.Text;

                    spImportTable.Rows.Add(newRow);
                    spImport.ItemsSource = spImportTable.DefaultView;
                    spImport.CanUserAddRows = false;

                    spimpid.Text = string.Empty;
                    spimpdes.Text = string.Empty;
                    spimpmd.Text = string.Empty;
                    spimppn.Text = string.Empty;
                    spimpmanu.Text = string.Empty;
                    spimptpu.Text = string.Empty;
                    spimploc.Text = string.Empty;
                    spimpmc.Text = string.Empty;
                    spstrloc.Text = string.Empty;
                    spstrlocno.Text = string.Empty;
                    spimpquan.Text = string.Empty;
                    spimpsaty.Text = string.Empty;
                    spimpunit.Text = string.Empty;
                    spimpprice.Text = string.Empty;
                    spimpcurrency.Text = string.Empty;
                    spdate.Text = string.Empty;
                    sptmpnote.Text = string.Empty;
                    spnewimg.Text = string.Empty;
                    spnewquote.Text = string.Empty;
                }
                else
                {
                    MessageBox.Show("Item ID is already in the table. Duplicate Item IDs are not allowed.");
                }
            }
        }

        private DataTable SPimrTable()
        {
            DataTable SPIMRTable = new DataTable();
            SPIMRTable.Columns.Add("ItemID", typeof(string));
            SPIMRTable.Columns.Add("Importer", typeof(string));
            SPIMRTable.Columns.Add("Import_quantity", typeof(int));
            SPIMRTable.Columns.Add("Import_date", typeof(DateTime));
            SPIMRTable.Columns.Add("Import_reason", typeof(string));

            return SPIMRTable;
        }

        private DataTable SPexpTable()
        {
            DataTable SPEXPTable = new DataTable();
            SPEXPTable.Columns.Add("ItemID", typeof(string));
            SPEXPTable.Columns.Add("Exporter", typeof(string));
            SPEXPTable.Columns.Add("Export_quantity", typeof(int));
            SPEXPTable.Columns.Add("Export_date", typeof(DateTime));
            SPEXPTable.Columns.Add("Export_reason", typeof(string));

            return SPEXPTable;
        }

        private void SpbtnImpAdd_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(spimrid.Text) ||
                string.IsNullOrEmpty(spimrimt.Text) ||
                string.IsNullOrEmpty(spimrquan.Text) ||
                !spimrdate.SelectedDate.HasValue ||
                string.IsNullOrEmpty(spimrrea.Text))
            {
                MessageBox.Show("Please fill in all fields before adding.");
                return;
            }

            SqlConnection sqlCon = ServerConnection.OpenConnection(PathReader.MNSP_link);
            {
                using (SqlCommand sqlCmd = new SqlCommand("SELECT Store_loc, Loc_no FROM Device_List WHERE ItemID = @ItemId", sqlCon))
                {
                    sqlCmd.Parameters.AddWithValue("@ItemId", spimrid.Text);
                    using (SqlDataReader reader = sqlCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            spimrstr.Text = reader["Store_loc"].ToString();
                            spimrloc.Text = reader["Loc_no"].ToString();
                        }
                        else
                        {
                            spimrstr.Text = "";
                            spimrloc.Text = "";
                            MessageBox.Show("Item ID not found in the database.");
                        }
                    }
                }

                using (SqlCommand sqlCmd = new SqlCommand("SELECT Image_name FROM Image_List WHERE ItemID = @ItemId", sqlCon))
                {
                    sqlCmd.Parameters.AddWithValue("@ItemId", spimrid.Text);
                    using (SqlDataReader reader = sqlCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string imagePath = Path.Combine(PathReader.Image_device, reader["Image_name"].ToString());

                            if (File.Exists(imagePath))
                            {
                                BitmapImage bitmapImage = new BitmapImage(new Uri(imagePath));
                                spimr_image.Source = bitmapImage;
                            }
                            else
                            {
                                spimr_image.Source = null;
                            }
                        }
                    }
                }
            }
            ServerConnection.CloseConnection(sqlCon);

            DataRow newRow = SPIMRTable.NewRow();
            newRow["ItemID"] = spimrid.Text;
            newRow["Importer"] = spimrimt.Text;
            newRow["Import_quantity"] = int.Parse(spimrquan.Text);
            newRow["Import_date"] = spimrdate.SelectedDate.Value;
            newRow["Import_reason"] = spimrrea.Text;
            SPIMRTable.Rows.Add(newRow);

            spimrGrid.ItemsSource = SPIMRTable.DefaultView;

            DataGridColumn importerColumn = spimrGrid.Columns.FirstOrDefault(col => col.Header.ToString() == "Importer");
            if (importerColumn != null)
            {
                importerColumn.IsReadOnly = true;
            }

            spimrid.Text = string.Empty;
            spimrimt.Text = string.Empty;
            spimrquan.Text = string.Empty;
            spimrdate.Text = string.Empty;
            spimrrea.Text = string.Empty;

            DataTable userTable = SQLDataTool.QueryUserData($"SELECT FullName FROM User_Registry WHERE UserID = '{UserLogTool.userInfor.Rows[0][0]}'", new List<SqlParameter>(), PathReader.DVM_link);
            spimrimt.Text = userTable.Rows[0]["FullName"].ToString();
        }

        private void SpbtnExpAdd_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(spexpid.Text) ||
                string.IsNullOrEmpty(spexpimt.Text) ||
                string.IsNullOrEmpty(spexpquan.Text) ||
                !spexpdate.SelectedDate.HasValue ||
                string.IsNullOrEmpty(spexprea.Text))
            {
                MessageBox.Show("Please fill in all fields before adding.");
                return;
            }

            SqlConnection sqlCon = ServerConnection.OpenConnection(PathReader.MNSP_link);
            {
                using (SqlCommand sqlCmd = new SqlCommand("SELECT Store_loc, Loc_no FROM Device_List WHERE ItemID = @ItemId", sqlCon))
                {
                    sqlCmd.Parameters.AddWithValue("@ItemId", spexpid.Text);
                    using (SqlDataReader reader = sqlCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            spexpstr.Text = reader["Store_loc"].ToString();
                            spexploc.Text = reader["Loc_no"].ToString();
                        }
                        else
                        {
                            spexpstr.Text = "";
                            spexploc.Text = "";
                            MessageBox.Show("Item ID not found in the database.");
                        }
                    }
                }

                using (SqlCommand sqlCmd = new SqlCommand("SELECT Image_name FROM Image_List WHERE ItemID = @ItemId", sqlCon))
                {
                    sqlCmd.Parameters.AddWithValue("@ItemId", spexpid.Text);
                    using (SqlDataReader reader = sqlCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string imagePath = Path.Combine(PathReader.Image_device, reader["Image_name"].ToString());

                            if (File.Exists(imagePath))
                            {
                                BitmapImage bitmapImage = new BitmapImage(new Uri(imagePath));
                                spexp_image.Source = bitmapImage;
                            }
                            else
                            {
                                spexp_image.Source = null;
                            }
                        }
                    }
                }
                ServerConnection.CloseConnection(sqlCon);
            }

            if (int.TryParse(spexpquan.Text, out int expValue) && int.TryParse(spde_quan.Text, out int deValue))
            {
                if (expValue > deValue)
                {
                    MessageBox.Show("Output quantity is incorrect.");
                }
                else
                {
                    DataRow newRow = SPEXPTable.NewRow();
                    newRow["ItemID"] = spexpid.Text;
                    newRow["Exporter"] = spexpimt.Text;
                    newRow["Export_quantity"] = int.Parse(spexpquan.Text);
                    newRow["Export_date"] = spexpdate.SelectedDate.Value;
                    newRow["Export_reason"] = spexprea.Text;
                    SPEXPTable.Rows.Add(newRow);

                    spexpGrid.ItemsSource = SPEXPTable.DefaultView;

                    DataGridColumn exporterColumn = spexpGrid.Columns.FirstOrDefault(col => col.Header.ToString() == "Exporter");
                    if (exporterColumn != null)
                    {
                        exporterColumn.IsReadOnly = true;
                    }

                    spexpid.Text = string.Empty;
                    spexpimt.Text = string.Empty;
                    spexpquan.Text = string.Empty;
                    spexpdate.Text = string.Empty;
                    spexprea.Text = string.Empty;

                    DataTable userTable = SQLDataTool.QueryUserData($"SELECT FullName FROM User_Registry WHERE UserID = '{UserLogTool.userInfor.Rows[0][0]}'", new List<SqlParameter>(), PathReader.DVM_link);
                    spexpimt.Text = userTable.Rows[0]["FullName"].ToString();
                }
            }
            else
            {
                MessageBox.Show("Invalid input. Please enter valid integers in both fields.");
            }
        }

        private void SpimrGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (spimrGrid.SelectedValue is DataRowView imr_selectedRow)
            {
                string selectedItemId = imr_selectedRow["ItemID"].ToString();
                Imr_GetDeviceAndImageInfo(selectedItemId);
            }
        }

        private void Imr_GetDeviceAndImageInfo(string itemId)
        {
            SqlConnection sqlCon = ServerConnection.OpenConnection(PathReader.MNSP_link);
            {
                using (SqlCommand sqlCmd = new SqlCommand("SELECT Store_loc, Loc_no FROM Device_List WHERE ItemID = @ItemId", sqlCon))
                {
                    sqlCmd.Parameters.AddWithValue("@ItemId", itemId);
                    using (SqlDataReader reader = sqlCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            spimrstr.Text = reader["Store_loc"].ToString();
                            spimrloc.Text = reader["Loc_no"].ToString();
                        }
                        else
                        {
                            spimrstr.Text = "";
                            spimrloc.Text = "";
                            MessageBox.Show("Item ID not found in the database.");
                        }
                    }
                }

                using (SqlCommand sqlCmd = new SqlCommand("SELECT Image_name FROM Image_List WHERE ItemID = @ItemId", sqlCon))
                {
                    sqlCmd.Parameters.AddWithValue("@ItemId", itemId);
                    using (SqlDataReader reader = sqlCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string imagePath = Path.Combine(PathReader.MNSP_link, reader["Image_name"].ToString());

                            if (File.Exists(imagePath))
                            {
                                BitmapImage bitmapImage = new BitmapImage(new Uri(imagePath));
                                spimr_image.Source = bitmapImage;
                            }
                            else
                            {
                                spimr_image.Source = null;
                            }
                        }
                    }
                }
            }
            ServerConnection.CloseConnection(sqlCon);
        }

        private void SpexpGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (spexpGrid.SelectedValue is DataRowView exp_selectedRow)
            {
                string selectedItemId = exp_selectedRow["ItemID"].ToString();
                Exp_GetDeviceAndImageInfo(selectedItemId);
            }
        }

        private void Exp_GetDeviceAndImageInfo(string itemId)
        {
            SqlConnection sqlCon = ServerConnection.OpenConnection(PathReader.MNSP_link);
            {
                using (SqlCommand sqlCmd = new SqlCommand("SELECT Store_loc, Loc_no FROM Device_List WHERE ItemID = @ItemId", sqlCon))
                {
                    sqlCmd.Parameters.AddWithValue("@ItemId", itemId);
                    using (SqlDataReader reader = sqlCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            spexpstr.Text = reader["Store_loc"].ToString();
                            spexploc.Text = reader["Loc_no"].ToString();
                        }
                        else
                        {
                            spexpstr.Text = "";
                            spexploc.Text = "";
                            MessageBox.Show("Item ID not found in the database.");
                        }
                    }
                }

                using (SqlCommand sqlCmd = new SqlCommand("SELECT Image_name FROM Image_List WHERE ItemID = @ItemId", sqlCon))
                {
                    sqlCmd.Parameters.AddWithValue("@ItemId", itemId);
                    using (SqlDataReader reader = sqlCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string imagePath = Path.Combine(PathReader.MNSP_link, reader["Image_name"].ToString());

                            if (File.Exists(imagePath))
                            {
                                BitmapImage bitmapImage = new BitmapImage(new Uri(imagePath));
                                spexp_image.Source = bitmapImage;
                            }
                            else
                            {
                                spexp_image.Source = null;
                            }
                        }
                    }
                }
            }
        }

        private void SpRegister_Click(object sender, RoutedEventArgs e)
        {
            if (spImportTable == null || spImportTable.Rows.Count == 0)
            {
                MessageBox.Show("No data to register.");
                return;
            }

            try
            {
                foreach (DataRow row in spImportTable.Rows)
                {
                    string insertQuery = "INSERT INTO Device_List (ItemID,Description,Model,Part_no,Manufacturer,TPU,Location,Machine,Store_loc,Loc_no,Quantity,Safety_stock,Unit ,Price,Currency,Register_date,Register,Noted) " +
                        "VALUES (@ItemID,@Description,@Model,@Part_no,@Manufacturer,@TPU,@Location,@Machine,@Store_loc,@Loc_no,@Quantity,@Safety_stock,@Unit ,@Price,@Currency,@Register_date,@Register,@Noted)";
                    List<SqlParameter> parameters = new List<SqlParameter>
                    {
                        new SqlParameter("@ItemID", row["ItemID"]),
                        new SqlParameter("@Description", SqlDbType.NVarChar) { Value = row["Description"] },
                        new SqlParameter("@Model", row["Model"]),
                        new SqlParameter("@Part_no", row["Part_no"]),
                        new SqlParameter("@Manufacturer", SqlDbType.NVarChar) { Value = row["Manufacturer"] },
                        new SqlParameter("@TPU", row["TPU"]),
                        new SqlParameter("@Location", row["Location"]),
                        new SqlParameter("@Machine", row["Machine"]),
                        new SqlParameter("@Store_loc", row["Store_loc"]),
                        new SqlParameter("@Loc_no", row["Loc_no"]),
                        new SqlParameter("@Quantity", SqlDbType.Int) { Value = row["Quantity"] },
                        new SqlParameter("@Safety_stock", SqlDbType.Int) { Value = row["Safety_stock"] },
                        new SqlParameter("@Unit", row["Unit"]),
                        new SqlParameter("@Price", SqlDbType.Decimal) { Value = row["Price"] },
                        new SqlParameter("@Currency", row["Currency"]),
                        new SqlParameter("@Register_date", SqlDbType.Date) { Value = row["Register_date"] },
                        new SqlParameter("@Register", SqlDbType.NVarChar) { Value = row["Register"] },
                        new SqlParameter("@Noted", SqlDbType.NVarChar) { Value = row["Noted"] },
                    };
                    SQLDataTool.ExecuteNonQuery(insertQuery, parameters, PathReader.MNSP_link);

                    spnewintroimg.Source = null;
                    string imagePath = row["Image_name"].ToString();
                    int lastIndex = imagePath.LastIndexOf('\\');
                    string imageName = imagePath.Substring(lastIndex + 1);

                    List<SqlParameter> imageUpdateParameters = new List<SqlParameter>
                    {
                        new SqlParameter("@ItemID", row["ItemID"]),
                        new SqlParameter("@Image_name", imageName)
                    };
                    SQLDataTool.ExecuteNonQuery("INSERT INTO Image_List (ItemID,Image_name) VALUES (@ItemID,@Image_name)", imageUpdateParameters, PathReader.MNSP_link);

                    File.Copy(imagePath, Path.Combine(PathReader.Image_device, imageName), true);
                    sppdfViewer.Source = null;
                    string pdfPath = row["Quotation_name"].ToString();
                    if (!string.IsNullOrEmpty(pdfPath) && File.Exists(pdfPath))
                    {
                        int pdflastIndex = pdfPath.LastIndexOf('\\');
                        string pdfName = pdfPath.Substring(pdflastIndex + 1);

                        List<SqlParameter> quoteUpdateParameters = new List<SqlParameter>
                        {
                            new SqlParameter("@ItemID", row["ItemID"]),
                            new SqlParameter("@Quotation_name", pdfName)
                        };
                        SQLDataTool.ExecuteNonQuery("INSERT INTO Quote_List (ItemID,Quotation_name) VALUES (@ItemID,@Quotation_name)", quoteUpdateParameters, PathReader.MNSP_link);

                        File.Copy(pdfPath, Path.Combine(PathReader.Quote_folder, pdfName), true);
                    }
                    spImportTable.Clear();
                }

                MessageBox.Show("Data inserted successfully.");
                LoadItemID();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inserting data: " + ex.Message);
            }
        }

        private void BtnspDel_Click(object sender, RoutedEventArgs e)
        {
            if (spImportTable != null && spImportTable.Rows.Count > 0 && spImport.SelectedValue != null)
            {
                DataRowView selectedRowView = (DataRowView)spImport.SelectedValue;
                DataRow selectedRow = selectedRowView.Row;
                spImportTable.Rows.Remove(selectedRow);
                spImport.ItemsSource = spImportTable.DefaultView;
            }
            else
            {
                MessageBox.Show("The table is empty.");
            }
        }

        private void Btnspimage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog imageFileDialog = new OpenFileDialog
            {
                Filter = "Image Files (*.jpg, *.jpeg, *.png)|*.jpg;*.jpeg;*.png",
                Title = "Choose Image"
            };

            bool? result = imageFileDialog.ShowDialog();
            if (result.HasValue && result.Value)
            {
                spnewimg.Text = imageFileDialog.FileName;
                string imagePath = imageFileDialog.FileName;

                if (!string.IsNullOrWhiteSpace(imagePath))
                {
                    BitmapImage bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.UriSource = new Uri(imagePath, UriKind.Absolute);
                    bitmapImage.EndInit();

                    spnewintroimg.Source = bitmapImage;
                }
            }
        }

        private void Btnspquote_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "PDF Files (*.pdf)|*.pdf|Word Documents (*.docx)|*.docx|Outlook Messages (*.msg)|*.msg|Image Files (*.jpg)|*.jpg"
            };

            bool? result = fileDialog.ShowDialog();

            if (result.HasValue && result.Value)
            {
                spnewquote.Text = fileDialog.FileName;
                sppdfViewer.Navigate(new Uri(fileDialog.FileName));
            }
        }

        private void Spimrimp_Click(object sender, RoutedEventArgs e)
        {
            if (SPIMRTable.Rows.Count == 0)
            {
                MessageBox.Show("Import Table is empty. No data to import.");
                return;
            }

            SqlConnection sqlCon = ServerConnection.OpenConnection(PathReader.MNSP_link);
            {
                using (SqlTransaction transaction = sqlCon.BeginTransaction())
                {
                    try
                    {
                        foreach (DataRow row in SPIMRTable.Rows)
                        {
                            string insertQuery = "INSERT INTO Import_History (ItemID, Importer, Import_quantity, Import_date, Import_reason) VALUES (@ItemID, @Importer, @Import_quantity, @Import_date, @Import_reason)";
                            List<SqlParameter> parameters = new List<SqlParameter>
                            {
                                new SqlParameter("@ItemID", SqlDbType.VarChar) { Value = row["ItemID"] },
                                new SqlParameter("@Importer", SqlDbType.NVarChar) { Value = row["Importer"] },
                                new SqlParameter("@Import_quantity", SqlDbType.Int) { Value = row["Import_quantity"] },
                                new SqlParameter("@Import_date", SqlDbType.Date) { Value = row["Import_date"] },
                                new SqlParameter("@Import_reason", SqlDbType.NVarChar) { Value = row["Import_reason"] }
                            };
                            SQLDataTool.ExecuteNonQuery(insertQuery, parameters, PathReader.MNSP_link);
                        }
                        transaction.Commit();
                        MessageBox.Show("Data imported successfully.");

                        OutlookTool.SendEmailWithExcelAttachment(SPIMRTable);
                        SPIMRTable.Clear();
                        spimrGrid.ItemsSource = SPIMRTable.DefaultView;
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show("Error while importing data: " + ex.Message);
                    }
                }
            }
        }

        private void Spexpid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            spexpquan.IsReadOnly = false;
            if (spexpid.SelectedValue != null)
            {
                string selectedItem = spexpid.SelectedValue.ToString();
                DataTable Quan_Table = SQLDataTool.QueryUserData($"SELECT Quantity FROM Device_List WHERE ItemID = '{selectedItem}'", new List<SqlParameter>(), PathReader.MNSP_link);
                spde_quan.Text = Quan_Table.Rows[0]["Quantity"].ToString();
                if (spde_quan.Text == "0")
                {
                    spexpquan.IsReadOnly = true;
                    MessageBox.Show("This device is temporarily unavailable, please top-up more.");
                }
            }
        }

        private void Spexpimp_Click(object sender, RoutedEventArgs e)
        {
            if (SPEXPTable.Rows.Count == 0)
            {
                MessageBox.Show("Export Table is empty. No data to export.");
                return;
            }
            SqlConnection sqlCon = ServerConnection.OpenConnection(PathReader.MNSP_link);
            {
                using (SqlTransaction transaction = sqlCon.BeginTransaction())
                {
                    try
                    {
                        foreach (DataRow row in SPEXPTable.Rows)
                        {
                            string insertQuery = "INSERT INTO Export_History (ItemID, Exporter, Export_quantity, Export_date, Export_reason) VALUES (@ItemID, @Exporter, @Export_quantity, @Export_date, @Export_reason)";
                            List<SqlParameter> parameters = new List<SqlParameter>
                            {
                                new SqlParameter("@ItemID", SqlDbType.VarChar) { Value = row["ItemID"] },
                                new SqlParameter("@Exporter", SqlDbType.NVarChar) { Value = row["Exporter"] },
                                new SqlParameter("@Export_quantity", SqlDbType.Int) { Value = row["Export_quantity"] },
                                new SqlParameter("@Export_date", SqlDbType.Date) { Value = row["Export_date"] },
                                new SqlParameter("@Export_reason", SqlDbType.NVarChar) { Value = row["Export_reason"] }
                            };
                            SQLDataTool.ExecuteNonQuery(insertQuery, parameters, PathReader.MNSP_link);
                        }
                        transaction.Commit();
                        MessageBox.Show("Data Exported successfully.");

                        OutlookTool.SendEmailWithExcelAttachment(SPEXPTable);
                        SPEXPTable.Clear();
                        spexpGrid.ItemsSource = SPEXPTable.DefaultView;
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show("Error while exporting data: " + ex.Message);
                    }
                }
            }
        }

        private void NumericTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            foreach (char c in e.Text)
            {
                if (!char.IsDigit(c))
                {
                    e.Handled = true;
                    break;
                }
            }
        }

        private void LoadChartHistoryData(string selectedRow)
        {
            DataTable exportData = SQLDataTool.QueryUserData("SELECT ItemID, Export_quantity, Export_date FROM Export_History WHERE ItemID = @selectedRow", new List<SqlParameter> { new SqlParameter("@selectedRow", selectedRow) }, PathReader.MNSP_link);
            DataTable importData = SQLDataTool.QueryUserData("SELECT ItemID, Import_quantity, Import_date FROM Import_History WHERE ItemID = @selectedRow", new List<SqlParameter> { new SqlParameter("@selectedRow", selectedRow) }, PathReader.MNSP_link);

            List<KeyValuePair<string, int>> chartData = new List<KeyValuePair<string, int>>();
            foreach (DataRow row in importData.Rows)
            {
                DateTime date = (DateTime)row["Import_date"];
                int quantity = (int)row["Import_quantity"];
                string value = "Import\n" + date.ToShortDateString();
                chartData.Add(new KeyValuePair<string, int>(value, quantity));
            }

            foreach (DataRow row in exportData.Rows)
            {
                DateTime date = (DateTime)row["Export_date"];
                int quantity = -(int)row["Export_quantity"];
                string value = "Export\n" + date.ToShortDateString();
                chartData.Add(new KeyValuePair<string, int>(value, quantity));
            }

            chartData.Sort((pair1, pair2) => DateTime.Parse(pair1.Key.Split('\n')[1]).CompareTo(DateTime.Parse(pair2.Key.Split('\n')[1])));
            ColumnSeries importExportSeries = new ColumnSeries
            {
                Title = "Item: " + selectedRow,
                FontSize = 13,
                Values = new ChartValues<int>(chartData.Select(dp => dp.Value)),
                DataLabels = true,
                Fill = Brushes.OrangeRed
            };

            spHistoryChart.AxisX.Clear();
            spHistoryChart.AxisX.Add(new Axis
            {
                Title = "",
                Labels = new List<string>(chartData.Select(dp => dp.Key)),
                LabelsRotation = 0,
                Foreground = Brushes.Black
            });

            spHistoryChart.AxisY.Clear();
            spHistoryChart.AxisY.Add(new Axis
            {
                Title = "Item: " + selectedRow,
                FontSize = 13,
                Foreground = Brushes.Black
            });

            spHistoryChart.Series.Clear();
            spHistoryChart.Series.Add(importExportSeries);
        }

        private void SpHisGridView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is DataGrid dataGrid)
            {
                if (dataGrid.SelectedValue is DataRowView selectedRowView)
                {
                    DataRow selectedRow = selectedRowView.Row;
                    string idItem = selectedRow["ItemID"].ToString();
                    ImageSource imageData = ImageViewer.GetImage(idItem, PathReader.MNSP_link, PathReader.Image_device);

                    sphis_image.Source = imageData ?? null;
                    LoadChartHistoryData(idItem);
                }
            }
        }

        private void ImportHistoryButton_Click(object sender, RoutedEventArgs e)
        {
            CheckSafetyStock();
            SPtempHistory = SQLDataTool.QueryUserData("SELECT * FROM Import_History", new List<SqlParameter>(), PathReader.MNSP_link);
            sphis_grid.ItemsSource = SPtempHistory.DefaultView;
            DateTimeFormat.DatetimeFormat(SPtempHistory, sphis_grid, "A");
            sphis_grid.CanUserAddRows = false;

            sphis_combobox.Items.Clear();
            foreach (DataColumn column in SPtempHistory.Columns)
            {
                sphis_combobox.Items.Add(column.ColumnName);
            }
        }

        private void ExportHistoryButton_Click(object sender, RoutedEventArgs e)
        {
            CheckSafetyStock();
            SPtempHistory = SQLDataTool.QueryUserData("SELECT * FROM Export_History", new List<SqlParameter>(), PathReader.MNSP_link);
            sphis_grid.ItemsSource = SPtempHistory.DefaultView;
            DateTimeFormat.DatetimeFormat(SPtempHistory, sphis_grid, "A");

            sphis_combobox.Items.Clear();
            foreach (DataColumn column in SPtempHistory.Columns)
            {
                sphis_combobox.Items.Add(column.ColumnName);
            }
        }

        private void CheckSafetyStock()
        {
            SqlConnection sqlCon = ServerConnection.OpenConnection(PathReader.MNSP_link);
            {
                using (SqlCommand sqlCmd = new SqlCommand("SELECT ItemID, Part_no, Description, Quantity, Safety_stock, TPU, Location FROM Device_List WHERE Quantity <= Safety_stock", sqlCon))
                {
                    using (SqlDataReader reader = sqlCmd.ExecuteReader())
                    {
                        DataTable dt = new DataTable();
                        dt.Columns.Add("ItemID", typeof(string));
                        dt.Columns.Add("Part_no", typeof(string));
                        dt.Columns.Add("Description", typeof(string));
                        dt.Columns.Add("Quantity", typeof(int));
                        dt.Columns.Add("Safety_stock", typeof(int));
                        dt.Columns.Add("TPU", typeof(string));
                        dt.Columns.Add("Location", typeof(string));

                        while (reader.Read())
                        {
                            string itemID = reader["ItemID"].ToString();
                            string partNo = reader["Part_no"].ToString();
                            string description = reader["Description"].ToString();
                            int quantity = Convert.ToInt32(reader["Quantity"]);
                            int safetyStock = Convert.ToInt32(reader["Safety_stock"]);
                            string tpu = reader["TPU"].ToString();
                            string location = reader["Location"].ToString();

                            dt.Rows.Add(itemID, partNo, description, quantity, safetyStock, tpu, location);
                        }

                        spsafety_grid.ItemsSource = dt.DefaultView;
                        spsafety_grid.CanUserAddRows = false;
                        SPtopup = dt;
                    }
                }
                ServerConnection.CloseConnection(sqlCon);
            }
        }

        private void SpHistory_TextChanged(object sender, TextChangedEventArgs e)
        {
            string searchText = spHistory_Text.Text;
            DataView dataView = SPtempHistory.DefaultView;

            if (searchText != null)
            {
                string selectedColumn = sphis_combobox.SelectedValue as string;
                if (!string.IsNullOrWhiteSpace(selectedColumn))
                {
                    DataColumn column = dataView.Table.Columns[selectedColumn];

                    dataView.RowFilter = column.DataType == typeof(DateTime)
                        ? DateTime.TryParse(searchText, out DateTime searchDate) ? $"{selectedColumn} = #{searchDate.ToShortDateString()}#" : "1=0"
                        : column.DataType == typeof(int)
                            ? int.TryParse(searchText, out int searchNumber) ? $"{selectedColumn} = {searchNumber}" : "1=0"
                            : $"{selectedColumn} LIKE '%{searchText}%'";
                }
            }
            sphis_grid.ItemsSource = dataView;
        }

        private void Spprintexceltopup_Click(object sender, RoutedEventArgs e)
        {
            ExcelTool.ExportExcelWithDialog(SPtopup, "SP_Topup", "Spare_TopupData");
        }

        private void Spmailtopup_Click(object sender, RoutedEventArgs e)
        {
            OutlookTool.SendEmailWithExcelAttachment(SPtopup);
        }

        private void Btnsp_changeimage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Image Files (*.jpg; *.png; *.bmp)|*.jpg;*.png;*.bmp"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string imagePath = openFileDialog.FileName;
                    UpdateImageNameInDatabase(spID.Text.ToString(), Path.GetFileName(imagePath));
                    BitmapImage bitmapImage = new BitmapImage(new Uri(imagePath));
                    spDevice_image.Source = bitmapImage;

                    File.Copy(imagePath, Path.Combine(PathReader.Image_device, Path.GetFileName(imagePath)), true);

                    MessageBox.Show("Image changed successfully.");
                }
                catch
                {
                    MessageBox.Show("Image changed successfully.");
                }
            }
        }

        private void UpdateImageNameInDatabase(string itemId, string newImageName)
        {
            SqlConnection connection = ServerConnection.OpenConnection(PathReader.MNSP_link);
            {
                using (SqlCommand command = new SqlCommand("UPDATE Image_List SET Image_name = @NewImageName WHERE ItemID = @ItemID", connection))
                {
                    command.Parameters.AddWithValue("@ItemID", itemId);
                    command.Parameters.AddWithValue("@NewImageName", newImageName);
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}