using System.Data;
using System.Windows.Data;
using System.Windows.Controls;
using System;
using System.Windows;

namespace MnS.lib
{
    public static class DateTimeFormat
    {
        public static void DatetimeFormat(DataTable dataTable, DataGrid dataGrid, string type)
        {
            try
            {
                dataGrid.ItemsSource = dataTable.DefaultView;

                foreach (DataGridColumn column in dataGrid.Columns)
                {
                    string columnHeader = column.Header.ToString();
                    if (columnHeader.Contains("date") || columnHeader.Contains("Date") || columnHeader.Contains("FDAT") || columnHeader.Contains("TDAT") || columnHeader.Contains("GDT") || columnHeader.Contains("MDT"))
                    {
                        if (column is DataGridTextColumn textColumn && type == "A")
                        {
                            textColumn.Binding = new Binding(columnHeader) { StringFormat = "dd/MMM/yyyy" };
                        }
                        else if (column is DataGridTextColumn textColumn1 && type == "B")
                        {
                            textColumn1.Binding = new Binding(columnHeader) { StringFormat = "dd/MM/yyyy" };
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex);
            }
        }

        public static void DatetimeFormat_Table(DataTable dataTable, string type)
        {
            try
            {
                foreach (DataColumn column in dataTable.Columns)
                {
                    if (column.ColumnName != null && !string.IsNullOrEmpty(column.ColumnName))
                    {
                        string columnName = column.ColumnName.ToUpper();

                        switch (type)
                        {
                            case "A":
                                if (columnName.Contains("DATE") || columnName.Contains("FDAT") || columnName.Contains("TDAT") || columnName.Contains("DAT"))
                                {
                                    ApplyDateFormat(dataTable, column.ColumnName, "dd/MMM/yyyy");
                                }
                                break;
                            case "B":
                                if (columnName.Contains("DATE") || columnName.Contains("FDAT") || columnName.Contains("TDAT") || columnName.Contains("DAT"))
                                {
                                    ApplyDateFormat(dataTable, column.ColumnName, "dd/MM/yyyy");
                                }
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        private static void ApplyDateFormat(DataTable dataTable, string columnName, string dateFormat)
        {
            foreach (DataRow row in dataTable.Rows)
            {
                if (row[columnName] != null && row[columnName] != DBNull.Value)
                {
                    if (DateTime.TryParse(row[columnName].ToString(), out DateTime dateValue))
                    {
                        row[columnName] = dateValue.ToString(dateFormat);
                    }
                }
            }
        }
    }
}