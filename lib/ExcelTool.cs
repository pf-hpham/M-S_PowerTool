using System;
using System.Data;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using OfficeOpenXml;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml.Style;
using System.Linq;

namespace MnS.lib
{
    public static class ExcelTool
    {
        public static string excel_filePath;

        public static void ExportExcelWithPath(string filePath, DataTable dataTable)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    for (int col = 1; col <= dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[1, col].Value = dataTable.Columns[col - 1].ColumnName;
                    }
                    for (int row = 0; row < dataTable.Rows.Count; row++)
                    {
                        for (int col = 0; col < dataTable.Columns.Count; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                        }
                    }
                    FileInfo fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo);

                    MessageBox.Show("Excel file saved successfully!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving Excel file: " + ex.Message);
            }
        }

        public static void ExportExcelWithDialog(DataTable dataTable, string sheetname, string fileName)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true,
                    FileName = fileName
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetname);
                        for (int col = 1; col <= dataTable.Columns.Count; col++)
                        {
                            worksheet.Cells[1, col].Value = dataTable.Columns[col - 1].ColumnName;
                        }
                        for (int row = 0; row < dataTable.Rows.Count; row++)
                        {
                            for (int col = 0; col < dataTable.Columns.Count; col++)
                            {
                                worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                            }
                        }
                        FormatDateColumns(worksheet);
                        ExcelFormat(worksheet);
                        FileInfo fileInfo = new FileInfo(filePath);
                        package.SaveAs(fileInfo);
                        MessageBox.Show("Excel file saved successfully!");
                        Process.Start(filePath);
                    }
                }
                else
                {
                    MessageBox.Show("Export canceled.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving Excel file: " + ex.Message);
            }
        }

        public static void ExportExcelWithDialog_Package(ExcelPackage package)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    package.SaveAs(new FileInfo(filePath));
                    MessageBox.Show("Excel file saved successfully!");
                    Process.Start(filePath);
                }
                else
                {
                    MessageBox.Show("Export canceled.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving Excel file: " + ex.Message);
            }
        }

        public static void ExportExcelWithName(DataTable dataTable, string fileName)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true,
                    FileName = fileName
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                        for (int col = 1; col <= dataTable.Columns.Count; col++)
                        {
                            worksheet.Cells[1, col].Value = dataTable.Columns[col - 1].ColumnName;
                        }
                        for (int row = 0; row < dataTable.Rows.Count; row++)
                        {
                            for (int col = 0; col < dataTable.Columns.Count; col++)
                            {
                                worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                            }
                        }
                        FileInfo fileInfo = new FileInfo(filePath);
                        package.SaveAs(fileInfo);
                        excel_filePath = filePath;
                        MessageBox.Show("Excel file saved successfully!");
                    }
                }
                else
                {
                    MessageBox.Show("Export canceled.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving Excel file: " + ex.Message);
            }
        }

        public static void ExportExcelWithListTable(List<DataTable> listTables, string fileName)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true,
                FileName = fileName
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;

                using (ExcelPackage package = new ExcelPackage())
                {
                    foreach (DataTable dataTable in listTables)
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(dataTable.TableName);
                        for (int col = 1; col <= dataTable.Columns.Count; col++)
                        {
                            string columnName = dataTable.Columns[col - 1].ColumnName;
                            worksheet.Cells[1, col].Value = columnName;

                            if (columnName.Contains("Date"))
                            {
                                worksheet.Column(col).Style.Numberformat.Format = "dd-mmm-yyyy";
                            }
                        }

                        for (int row = 0; row < dataTable.Rows.Count; row++)
                        {
                            for (int col = 0; col < dataTable.Columns.Count; col++)
                            {
                                worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                            }
                        }

                        ExcelFormat(worksheet);
                    }

                    package.SaveAs(new FileInfo(filePath));
                    excel_filePath = filePath;
                    MessageBox.Show("Successfully exporting to Excel file.");
                    Process.Start(filePath);
                }
            }
        }

        public static void CreateExcelFileFromDataTable(DataTable dataTable, string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook excelWorkbook = null;
            Excel.Worksheet excelWorksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelWorkbook = excelApp.Workbooks.Add();
                excelWorksheet = excelWorkbook.Sheets[1];

                for (int col = 1; col <= dataTable.Columns.Count; col++)
                {
                    excelWorksheet.Cells[1, col] = dataTable.Columns[col - 1].ColumnName;
                }

                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        excelWorksheet.Cells[row + 2, col + 1] = dataTable.Rows[row][col].ToString();
                    }
                }
                excelWorkbook.SaveAs(filePath);
            }
            catch (Exception ex)
            {
                throw new Exception("Error exporting Excel file: " + ex.Message);
            }
            finally
            {
                if (excelWorkbook != null)
                {
                    excelWorkbook.Close();
                    Marshal.ReleaseComObject(excelWorkbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                if (excelWorksheet != null)
                {
                    Marshal.ReleaseComObject(excelWorksheet);
                }
            }
        }

        public static void ExcelFormat(ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension != null)
            {
                worksheet.Cells.AutoFitColumns();
                worksheet.Cells[worksheet.Dimension.Address].AutoFilter = true;

                var dataRange = worksheet.Cells[worksheet.Dimension.Address];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                var headerRow = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column];
                headerRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerRow.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkCyan);
                headerRow.Style.Font.Color.SetColor(System.Drawing.Color.White);
                headerRow.Style.Font.Bold = true;

                var resultColumn1 = worksheet.Cells[2, 5, worksheet.Dimension.End.Row, 5];
                var resultColumn2 = worksheet.Cells[2, 4, worksheet.Dimension.End.Row, 4];
                foreach (var cell in resultColumn1)
                {
                    if (cell.Text == "D")
                    {
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            var currentCell = worksheet.Cells[cell.Start.Row, col];
                            currentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            currentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSalmon);
                        }
                    }
                }
                foreach (var cell in resultColumn2)
                {
                    if (cell.Text == "D")
                    {
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            var currentCell = worksheet.Cells[cell.Start.Row, col];
                            currentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            currentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSalmon);
                        }
                    }
                }
            }
        }

        public static void FormatDateColumns(ExcelWorksheet worksheet)
        {
            foreach (var column in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column].Select((col, index) => new { col = col, index = index + 1 }))
            {
                if (column.col.Text.Contains("Date") || column.col.Text.Contains("date") || column.col.Text.Contains("DAT") || column.col.Text.Contains("GDT") || column.col.Text.Contains("MDT"))
                {
                    ExcelRange dateColumn = worksheet.Cells[2, column.index, worksheet.Dimension.End.Row, column.index];
                    dateColumn.Style.Numberformat.Format = "dd/MM/yyyy";
                }
            }
        }
    }
}