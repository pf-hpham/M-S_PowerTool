using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Controls;

namespace MnS.lib
{
    public static class ComboBoxTool
    {
        /// <summary>
        /// Create a list of choice for combobox <br></br> <br></br>
        /// Parameters: <br></br>
        /// "query" Query to SQL database <br></br>
        /// "parameters" List value of choice <br></br>
        /// "connection" Connectionstring <br></br>
        /// "comboBox" Combobox x:Name <br></br>
        /// "memberPath" <br></br>
        /// "valuePath" <br></br>
        /// "selectedIndex" First display value
        /// </summary>
        public static void LoadCombobox(string query, List<SqlParameter> parameters, string connection, ComboBox comboBox, string memberPath, string valuePath, int selectedIndex)
        {
            try
            {
                DataTable dt = SQLDataTool.QueryUserData(query, parameters, connection);

                comboBox.ItemsSource = dt.DefaultView;
                comboBox.DisplayMemberPath = memberPath;
                comboBox.SelectedValuePath = valuePath;
                comboBox.SelectedIndex = selectedIndex;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error loading ComboBox data: " + ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="comboBox"></param>
        public static void DisplayComboBox(DataTable dataTable, ComboBox comboBox)
        {
            comboBox.Items.Clear();
            foreach (DataColumn column in dataTable.Columns)
            {
                comboBox.Items.Add(column.ColumnName);
            }
        }
    }
}