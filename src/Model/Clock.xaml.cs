using Microsoft.Win32;
using MnS.lib;
using OfficeOpenXml;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace MnS
{
    public partial class Clock : Window
    {
        #region Variable
        private string filePath;
        private string week;
        private int wk_col;
        private ExcelWorksheet sl_worksheet;
        private DataTable dataTable;
        #endregion

        public Clock()
        {
            UserLogTool.UserData("Using Clock function");
            InitializeComponent();
            FillComboBox();
        }

        private void ClkOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (openFileDialog.ShowDialog() == true)
            {
                filePath = openFileDialog.FileName;
                clk_link.Text = filePath;
            }
            ReadExcelAndFillComboBox(filePath);
        }

        private void ReadExcelAndFillComboBox(string filePath)
        {
            clk_name.Items.Clear();
            clk_name.Items.Add("Choose Name");

            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Name.StartsWith("0") || worksheet.Name.StartsWith("1"))
                    {
                        object value = worksheet.Cells["C4"].Value;
                        if (value != null)
                        {
                            clk_name.Items.Add(value.ToString());
                        }
                    }
                }
            }
            clk_name.SelectedIndex = 0;
        }

        private void FillComboBox()
        {
            for (int i = 1; i <= 53; i++)
            {
                clk_week.Items.Add("Week: " + i);
            }
        }

        private void Update_View()
        {
            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                {
                    object value = worksheet.Cells["C4"].Value;
                    if (value != null && value.ToString() == clk_name.SelectedItem.ToString())
                    {
                        for (int col = 5; col <= 57; col++)
                        {
                            object cellValue = worksheet.Cells[7, col].Value;
                            week = "Week: " + cellValue;
                            if (cellValue != null && week == clk_week.SelectedItem.ToString())
                            {
                                object totalValue = worksheet.Cells[5, col].Value;
                                if (totalValue != null)
                                {
                                    clk_total.Text = totalValue.ToString();
                                    wk_col = col;
                                    sl_worksheet = worksheet;

                                    CreateDataTableWithWeekColumn(worksheet);
                                }
                            }
                        }
                        break;
                    }
                }
            }
        }

        private void Week_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (filePath != null)
            {
                if (clk_name.SelectedItem != null && clk_week.SelectedItem != null)
                {
                    Update_View();
                }
            }
        }

        private void CreateDataTableWithWeekColumn(ExcelWorksheet worksheet)
        {
            dataTable = new DataTable();

            dataTable.Columns.Add("No", typeof(string));
            dataTable.Columns.Add("Category", typeof(string));
            dataTable.Columns.Add("Incident No", typeof(string));
            dataTable.Columns.Add("CC", typeof(string));
            dataTable.Columns.Add("Hours", typeof(string));

            int rowCount = worksheet.Dimension.Rows;

            for (int row = 8; row <= rowCount; row++)
            {
                string no = worksheet.Cells[row, 1].Value?.ToString();
                string category = worksheet.Cells[row, 2].Value?.ToString();
                string incidentNo = worksheet.Cells[row, 3].Value?.ToString();
                string cc = worksheet.Cells[row, 4].Value?.ToString();
                string hours = worksheet.Cells[row, wk_col].Value?.ToString();

                if (!string.IsNullOrEmpty(hours))
                {
                    dataTable.Rows.Add(no, category, incidentNo, cc, hours);
                }
            }

            DateTimeFormat.DatetimeFormat(dataTable, clk_grid, "A");
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Please choose the Excel Clocking time first.");
                return;
            }

            bool foundEmptyRow = false;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Name == sl_worksheet.Name)
                    {
                        for (int row = 8; row <= worksheet.Dimension.End.Row; row++)
                        {
                            if (string.IsNullOrEmpty(worksheet.Cells[row, 3].Value?.ToString()))
                            {
                                foundEmptyRow = true;

                                string previousValue = worksheet.Cells[row - 1, 1].Value?.ToString();
                                int currentValue = string.IsNullOrEmpty(previousValue) ? 1 : int.Parse(previousValue) + 1;
                                string category = clk_cate.SelectedItem.ToString();
                                if (category.StartsWith("System.Windows.Controls.ComboBoxItem: "))
                                {
                                    category = category.Replace("System.Windows.Controls.ComboBoxItem: ", "");
                                }
                                string incidentNo = clk_incident.Text;
                                string cc = clk_cc.SelectedItem.ToString();
                                if (cc.StartsWith("System.Windows.Controls.ComboBoxItem: "))
                                {
                                    cc = cc.Replace("System.Windows.Controls.ComboBoxItem: ", "");
                                    cc = cc.Substring(0, 4);
                                }
                                string clockTime = clk_clocktime.Text;
                                clockTime = string.Format("{0:0.00}", double.Parse(clockTime));
                                int weekColumn = wk_col;

                                worksheet.Cells[row, 1].Value = currentValue;
                                worksheet.Cells[row, 2].Value = category;
                                worksheet.Cells[row, 3].Value = incidentNo;
                                worksheet.Cells[row, 4].Value = cc;
                                worksheet.Cells[row, weekColumn].Value = clockTime;

                                for (int col = 1; col <= weekColumn; col++)
                                {
                                    worksheet.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                }

                                break;
                            }
                        }

                        if (!foundEmptyRow)
                        {
                            MessageBox.Show("Cannot found the empty row to add new data.");
                        }
                        else
                        {
                            package.Save();
                            MessageBox.Show("Add new data successfully!");
                        }
                    }
                }
            }

            Update_View();
        }

        private void Insert_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            if (clk_grid.SelectedItem != null)
            {
                DataRowView selectedRow = (DataRowView)clk_grid.SelectedItem;

                string value1 = selectedRow["No"].ToString();
                string value2 = selectedRow["Category"].ToString();
                string value3 = selectedRow["Incident No"].ToString();

                RemoveRowFromExcel(value1, value2, value3);

                dataTable.Rows.Remove(selectedRow.Row);

                clk_grid.ItemsSource = dataTable.DefaultView;
            }
            else
            {
                MessageBox.Show("Please select a row to remove.", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void RemoveRowFromExcel(string value1, string value2, string value3)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    if (worksheet.Name == sl_worksheet.Name)
                    {
                        for (int row = 8; row <= worksheet.Dimension.End.Row; row++)
                        {
                            if (worksheet.Cells[row, 1].Value.ToString() == value1 &&
                                worksheet.Cells[row, 2].Value.ToString() == value2 &&
                                worksheet.Cells[row, 3].Value.ToString() == value3)
                            {
                                worksheet.DeleteRow(row);
                                break;
                            }
                        }
                        excelPackage.Save();
                    }
                }
            }
        }
    }
}