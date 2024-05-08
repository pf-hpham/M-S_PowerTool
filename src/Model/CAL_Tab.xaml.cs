using MnS.lib;
using System;
using System.IO;
using System.Linq;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows;
using System.Globalization;
using System.Data.SqlClient;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MnS
{
    public partial class CAL_Tab : UserControl
    {
        #region
        public static DataTable CALTable = new DataTable();
        private DataTable autolink_Table = new DataTable();
        private DataTable gridTable = new DataTable();
        private string temp_Path = null;
        private bool Flag = false;
        #endregion

        public CAL_Tab()
        {
            UserLogTool.UserData("Using Calibration function");
            InitializeComponent();
            CalYear();

            for (int i = 1; i <= 12; i++)
            {
                cal_month.Items.Add(DateTimeFormatInfo.CurrentInfo.GetAbbreviatedMonthName(i));
            }

            cal_year.Items.Add("2023");
            cal_year.Items.Add("2024");

            cal_month.SelectedIndex = -1;
            cal_year.SelectedIndex = -1;
        }

        #region All Button Click
        private async void Filterload_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                loadingProg.Visibility = Visibility.Visible;
                loadingProg.Value = 0;

                DataTable calTable = SQLDataTool.QueryUserData("SELECT * FROM Gage_Master", new List<SqlParameter>(), PathReader.CAL_link);
                CALTable = calTable;
                await ProgressBarTool.ProgressBarAsync(CALTable, (progress) =>
                {
                    ProgressBarTool.UpdateProgressBar(loadingProg, progress);
                });
                loadingProg.Visibility = Visibility.Hidden;

                DataContext = new CALModelView(calTable);
                MessageBox.Show("Loaded Database successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex);
            }
        }

        private void CALexport_Click(object sender, RoutedEventArgs e)
        {
            if (CALTable != null)
            {
                ExcelTool.ExportExcelWithDialog(CALTable, "CAL", "Calibration_database");
            }
            else
            {
                MessageBox.Show("None of Data to Export.");
            }
        }

        private void CALSend_Click(object sender, RoutedEventArgs e)
        {
            if (CALTable != null)
            {
                OutlookTool.SendEmailWithExcelAttachment(CALTable);
            }
            else
            {
                MessageBox.Show("None of Data to Send.");
            }
        }

        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            using (System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = folderBrowserDialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    string selectedFolder = folderBrowserDialog.SelectedPath;
                    scrcal_link.Text = selectedFolder;

                    DataTable pdfDataTable = new DataTable();
                    pdfDataTable.Columns.Add("file_name", typeof(string));
                    pdfDataTable.Columns.Add("save_name", typeof(string));
                    pdfDataTable.Columns.Add("source_link", typeof(string));
                    pdfDataTable.Columns.Add("save_link", typeof(string));

                    string[] pdfFiles = Directory.GetFiles(selectedFolder, "*.pdf");

                    foreach (string pdfFile in pdfFiles)
                    {
                        string fileName = Path.GetFileName(pdfFile);
                        string saveName = fileName;
                        string pdfLink = pdfFile;
                        string saveLink = "";

                        pdfDataTable.Rows.Add(fileName, saveName, pdfLink, saveLink);
                    }
                    DateTimeFormat.DatetimeFormat(pdfDataTable, pdf_grid, "A");
                    pdf_grid.CanUserAddRows = false;
                    autolink_Table = pdfDataTable;

                    DataGridColumn pdfLinkColumn = pdf_grid.Columns.FirstOrDefault(c => c.Header.ToString() == "source_link");
                    DataGridColumn fileNameColumn = pdf_grid.Columns.FirstOrDefault(c => c.Header.ToString() == "file_name");
                    Flag = true;
                    if (pdfLinkColumn != null)
                    {
                        pdfLinkColumn.Visibility = Visibility.Collapsed;
                        fileNameColumn.IsReadOnly = true;
                    }
                }
            }
        }

        private void Createlink_Click(object sender, RoutedEventArgs e)
        {
            if (Flag == true)
            {
                DataView view = (DataView)pdf_grid.ItemsSource;
                DataTable dataTable = view.Table.Clone();
                foreach (DataRowView row in view)
                {
                    string fileName = row["file_name"].ToString();
                    string saveName = row["save_name"].ToString();
                    string source_link = scrcal_link.Text + "\\" + fileName;
                    string save_link = scrcal_link.Text + "\\" + saveName;

                    try
                    {
                        File.Move(source_link, save_link);
                        pdf_show.Navigate("about:blank");
                        Directory.Delete(temp_Path, true);
                    }
                    catch { }

                    row["file_name"] = row["save_name"];
                    row["source_link"] = scrcal_link.Text + "\\" + saveName;

                    stt_label.Text = "load..." + saveName;

                    string[] part_a = saveName.Split('-');
                    string folder = part_a[0].Trim();

                    string[] part_b = saveName.Split(' ');
                    string folder_1 = part_b[0].Trim();

                    string[] subfolders = Directory.GetDirectories(PathReader.CALrecord_file);
                    foreach (string subfolder in subfolders)
                    {
                        if (subfolder.Contains(folder))
                        {
                            Path.GetFileName(subfolder);
                            folder = subfolder;
                            break;
                        }
                    }

                    row["save_link"] = folder + "\\" + folder_1 + "\\" + saveName;
                    // row["save_link"] = "C:\\Users\\hpham\\Desktop\\New folder(4)\\" + folder_1 + "\\" + saveName;
                
                    dataTable.ImportRow(row.Row);
                }
                autolink_Table = dataTable;
                DateTimeFormat.DatetimeFormat(autolink_Table, pdf_grid, "A");
                DataGridColumn pdfLinkColumn = pdf_grid.Columns.FirstOrDefault(c => c.Header.ToString() == "source_link");
                pdfLinkColumn.Visibility = Visibility.Collapsed;
                stt_label.Text = "created save-link successfully.";
                Flag = false;
            }
            else
            {
                MessageBox.Show("Please choose Calib records folder.");
            }
        }

        private void MSPsearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string selectedYear = cal_year.SelectedItem?.ToString();
                string selectedMonth = cal_month.SelectedItem?.ToString();

                if (string.IsNullOrEmpty(selectedYear) || string.IsNullOrEmpty(selectedMonth))
                {
                    MessageBox.Show("Please select a year and a month.");
                    return;
                }

                string excelFilePath = Path.Combine(PathReader.CALmsp_file, $"{selectedYear}\\Calibration_{selectedYear}.xlsx");

                if (!File.Exists(excelFilePath))
                {
                    MessageBox.Show("Excel file not found.");
                    return;
                }

                DataTable dt = new DataTable();

                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.FirstOrDefault(part => part.Uri.OriginalString.EndsWith($"{selectedMonth}.xml"));

                    if (worksheetPart == null)
                    {
                        MessageBox.Show($"Worksheet for {selectedMonth} not found in the Excel file.");
                        return;
                    }

                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        if (row.RowIndex == 6)
                        {
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                dt.Columns.Add(cell.CellValue.Text);
                            }
                        }
                        else if (row.RowIndex > 6)
                        {
                            DataRow newRow = dt.Rows.Add();
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                newRow[cell.CellReference.InnerText.ToCharArray()[0] - 'A'] = GetCellValue(cell, workbookPart);
                            }
                        }
                    }
                }

                int countTamsui = 0;
                int countTechmaster = 0;
                gridTable = dt;

                foreach (DataRow row in dt.Rows)
                {
                    string calibratorValue = row["Calibrated_By"].ToString();

                    if (calibratorValue.Equals("TAMSUI", StringComparison.OrdinalIgnoreCase))
                    {
                        countTamsui++;
                    }
                    else if (calibratorValue.Equals("TECHMASTER", StringComparison.OrdinalIgnoreCase))
                    {
                        countTechmaster++;
                    }
                }

                nmbTamsui.Content = countTamsui.ToString();
                nmbTech.Content = countTechmaster.ToString();

                cal_excel.ItemsSource = dt.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            SharedStringTablePart stringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (stringTablePart != null)
            {
                SharedStringItem sharedStringItem = stringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue.Text));
                return sharedStringItem.Text.Text;
            }
            else
            {
                return cell.CellValue.Text;
            }
        }

        private void ShowRecord_Click(object sender, RoutedEventArgs e)
        {
            DataTable recordData = SQLDataTool.QueryUserData("SELECT * FROM Calib_Attachments", new List<SqlParameter>(), PathReader.CAL_link);
            string cal_jigName = cal_jigname.Text.ToString();
            int cal_rctimeToFind = (int)cal_rctime.SelectedValue;
            DataRow[] filteredRows = recordData.Select($"Gage_ID = '{cal_jigName}'");
            List<DataRow> rowsToDisplay = new List<DataRow>();

            foreach (DataRow row in filteredRows)
            {
                if (row["Calibration_Date"] is DateTime calibrationDate)
                {
                    if (calibrationDate.Year == cal_rctimeToFind)
                    {
                        rowsToDisplay.Add(row);
                    }
                }
            }
            if (rowsToDisplay.Count > 0)
            {
                string filePath = Regex.Replace(rowsToDisplay[0]["AttachPath"].ToString(), "^[A-Z]:\\\\", "\\\\pfvn-netapp1\\files\\");
                cal_rcView.Navigate(new Uri(filePath));
            }
            else
            {
                MessageBox.Show("No matching data found.");
            }
        }

        private void ExportMSP_Click(object sender, RoutedEventArgs e)
        {

            if (gridTable != null)
            {
                ExcelTool.ExportExcelWithDialog(gridTable, "Exp_MSP", "MSP_Data");
            }
            else
            {
                MessageBox.Show("None of Data to Export.");
            }
        }

        private void MSPSend_Click(object sender, RoutedEventArgs e)
        {
            if (gridTable != null)
            {
                OutlookTool.SendEmailWithExcelAttachment(gridTable);
            }
            else
            {
                MessageBox.Show("None of Data to Send.");
            }
        }
        #endregion

        private void Pdfshow_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (pdf_grid.SelectedItem != null)
                {
                    DataRowView selectedRow = (DataRowView)pdf_grid.SelectedItem;
                    string sourceLink = selectedRow["source_link"].ToString();
                    temp_Path = scrcal_link.Text + "\\temp";
                    string tempFilePath = Path.Combine(temp_Path, selectedRow["file_name"].ToString());

                    Directory.CreateDirectory(temp_Path);
                    File.Copy(sourceLink, tempFilePath, true);
                    pdf_show.Navigate(new Uri(tempFilePath));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CalYear()
        {
            int currentYear = DateTime.Now.Year;
            for (int year = 2017; year <= currentYear; year++)
            {
                cal_rctime.Items.Add(year);
            }
        }

        private void MSPRefresh_Click(object sender, RoutedEventArgs e)
        {
            string owner = "";
            if (callocmsp_src.SelectedIndex == 0)
            {
                owner = "TM";
            }
            else if (callocmsp_src.SelectedIndex == 1)
            {
                owner = "P&E";
            }
            DataTable calData = SQLDataTool.QueryUserData($"SELECT Gage_ID, Gage_SN, Model_No, Manufacturer, GM_Owner, Description, Current_Location, Calibrator, Calibration_Frequency, Calibration_Frequency_UOM, Next_Due_Date, Last_Calibration_Date, Status, User_Defined, Calibrated_By FROM Gage_Master WHERE (Calibrator LIKE '%TAMSUI%' OR Calibrator LIKE '%TECHMASTER%' OR Calibrated_By LIKE '%TAMSUI%' OR Calibrated_By LIKE '%TECHMASTER%') AND (Status = 1 OR Status = 4) AND User_Defined LIKE '%EXTERNAL%' AND GM_Owner LIKE '%{owner}%'", new List<SqlParameter>(), PathReader.CAL_link);
            calData.Columns["Gage_ID"].ColumnName = "ID";
            calData.Columns["Gage_SN"].ColumnName = "SN";
            calData.Columns["Model_No"].ColumnName = "Model";
            calData.Columns["GM_Owner"].ColumnName = "Owner";
            calData.Columns["Description"].ColumnName = "Description";
            calData.Columns["Current_Location"].ColumnName = "Location";
            calData.Columns["Calibration_Frequency"].ColumnName = "Frequency";
            calData.Columns["Calibration_Frequency_UOM"].ColumnName = "UOM";
            calData.Columns["Next_Due_Date"].ColumnName = "Due Date";
            calData.Columns["Last_Calibration_Date"].ColumnName = "Calib Date";
            calData.Columns["User_Defined"].ColumnName = "Method";
            calData.Columns["Calibrated_By"].ColumnName = "Calibrated By";

            DataColumn notesColumn = new DataColumn("Notes", typeof(string));
            DataColumn calibColumn = new DataColumn("Calib date", typeof(string));
            int statusColumnIndex = calData.Columns["Status"].Ordinal;
            int notesColumnIndex = statusColumnIndex;
            calData.Columns.Add(notesColumn);
            calData.Columns.Add(calibColumn);
            notesColumn.SetOrdinal(notesColumnIndex);

            DataView[] monthDataViews = new DataView[12];
            List<DataTable> dataTables = new List<DataTable>();

            for (int i = 1; i < 13; i++)
            {
                DateTime startOfMonth = new DateTime();
                DateTime endOfMonth = new DateTime();

                if (i == 1)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 4);
                    endOfMonth = new DateTime(DateTime.Now.Year, i + 2, 7);
                }
                else if (i == 2)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 4);
                    endOfMonth = new DateTime(DateTime.Now.Year, i + 1, 4);
                }
                else if (i == 3)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 7);
                    endOfMonth = new DateTime(DateTime.Now.Year, i + 1, 4);
                }
                else if (i == 4)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 4);
                    endOfMonth = new DateTime(DateTime.Now.Year, i + 1, 9);
                }
                else if (i == 5)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 4);
                    endOfMonth = new DateTime(DateTime.Now.Year, i + 1, 6);
                }
                else if (i == 6)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 7);
                    endOfMonth = new DateTime(DateTime.Now.Year, i + 1, 4);
                }
                else if (i == 7)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 5);
                    endOfMonth = new DateTime(DateTime.Now.Year, i + 1, 8);
                }
                else if (i == 8)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 9);
                    endOfMonth = new DateTime(DateTime.Now.Year, i + 1, 12);
                }
                else if (i == 9)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 13);
                    endOfMonth = new DateTime(DateTime.Now.Year, i + 1, 11);
                }
                else if (i == 10)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 12);
                    endOfMonth = new DateTime(DateTime.Now.Year, i + 1, 7);
                }
                else if (i == 11)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 8);
                    endOfMonth = new DateTime(DateTime.Now.Year, i + 1, 7);
                }
                else if (i == 12)
                {
                    startOfMonth = new DateTime(DateTime.Now.Year, i, 7);
                    endOfMonth = new DateTime(DateTime.Now.Year + 1, 1, 4);
                }

                string filterExpression = $"[Due Date] >= #{startOfMonth.ToString("MM/dd/yyyy")}# AND [Due Date] <= #{endOfMonth.ToString("MM/dd/yyyy")}#";
                monthDataViews[i-1] = new DataView(calData, filterExpression, "", DataViewRowState.CurrentRows);

                DataTable monthDataTable = monthDataViews[i - 1].ToTable();
                monthDataTable.TableName = $"Month_{i}";
                dataTables.Add(monthDataTable);
            }

            ExcelTool.ExportExcelWithListTable(dataTables, "CAL_2024");
            cal_excel.ItemsSource = calData.DefaultView;
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            foreach (DataRow row in autolink_Table.Rows)
            {
                string sourceLink = row["source_link"].ToString();
                string saveLink = row["save_link"].ToString();
                try
                {
                    if (!Directory.Exists(Path.GetDirectoryName(saveLink)))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(saveLink));
                    }

                    File.Copy(sourceLink, saveLink, true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            MessageBox.Show("All records copied successfully!");
        }
    }
}