using MnS.lib;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
namespace MnS
{
    public partial class Movex : Window
    {
        #region
        private OdbcConnection mv_cnt;
        public delegate void UpdateProgressDelegate(int value, string prog);

        public static DataTable rt_table = new DataTable();
        public static DataTable bm_table = new DataTable();
        public static DataTable manual_table = new DataTable();
        public static DataTable wh_doctable = new DataTable();
        public static DataTable wh_instructiontable = new DataTable();
        public static DataTable wh_alternatetable = new DataTable();
        public static DataTable wh_materialtable = new DataTable();
        public static DataTable wh_depttable = new DataTable();
        public static DataTable ecn_table = new DataTable();

        public string item_input;
        public List<string> itemlist_input;
        public string item_type;
        public string wh_inputmanual;
        public string wh_inputdoc;
        public string wh_instructioninput;
        public string wh_alterinput;
        public string wh_materialinput;
        public string wh_deptinput;
        public string ecn_input;
        public string item_No;
        public string alter;

        public string DSN;
        public string CommandText_1;
        public string CommandText_2;
        public string CommandText_3;
        public string CommandText_4;
        public string CommandText_5;
        public string CommandText_6;
        public string CommandText_7;
        public string CommandText_8;
        public string CommandText_8A;
        public string CommandText_9;
        public string CommandText_10;
        public string CommandText_11;
        public string CommandText_12;

        public bool expand_wh;
        public bool expand_whType;
        public bool expand_whch;
        public bool expand_ecn;
        #endregion

        public Movex()
        {
            UserLogTool.UserData("Using Movex function");
            InitializeComponent();
            ReadCommandLine();
            expand_wh = false;
            expand_whType = false;
            expand_whch = false;
            expand_ecn = false;
            rtInput.Focus();
            rtSTD.IsChecked = true;
        }

        private void ReadCommandLine()
        {
            string[] lines = File.ReadAllLines(PathReader.Movex);
            foreach (string line in lines)
            {
                if (line.Contains("CommandText"))
                {
                    string[] part = line.Split(new char[] { '=' }, 2);
                    if (part[0].ToString() == "CommandText_1")
                    {
                        CommandText_1 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_2")
                    {
                        CommandText_2 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_3")
                    {
                        CommandText_3 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_4")
                    {
                        CommandText_4 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_5")
                    {
                        CommandText_5 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_6")
                    {
                        CommandText_6 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_7")
                    {
                        CommandText_7 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_8")
                    {
                        CommandText_8 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_8A")
                    {
                        CommandText_8A = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_9")
                    {
                        CommandText_9 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_10")
                    {
                        CommandText_10 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_11")
                    {
                        CommandText_11 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_12")
                    {
                        CommandText_12 = part[1].ToString();
                    }
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

        private void MnInput_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && item_no.SelectedIndex == 1)
            {
                e.Handled = true;
                string clipboardText = Clipboard.GetText();
                string[] lines = clipboardText.Split(new string[] { Environment.NewLine, "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries);
                rtInput.Text = string.Join(", ", lines);
            }
        }

        private void RtExport_tb_DropDownOpened(object sender, EventArgs e)
        {
            rtExport_tb.Items.Clear();
            if (rt_table != null && rt_table.Rows.Count > 0)
            {
                ComboBoxItem rtItem = new ComboBoxItem
                {
                    Content = "Routing_" + item_input
                };
                rtExport_tb.Items.Add(rtItem);
            }

            if (bm_table != null && bm_table.Rows.Count > 0)
            {
                ComboBoxItem bmItem = new ComboBoxItem
                {
                    Content = "Bom_" + item_input
                };
                rtExport_tb.Items.Add(bmItem);
            }

            if (manual_table != null && manual_table.Rows.Count > 0)
            {
                ComboBoxItem whItem = new ComboBoxItem
                {
                    Content = "Where-used_" + wh_inputmanual
                };
                rtExport_tb.Items.Add(whItem);
            }

            if (wh_doctable != null && wh_doctable.Rows.Count > 0)
            {
                ComboBoxItem whItem = new ComboBoxItem
                {
                    Content = "Where-used_" + wh_inputdoc
                };
                rtExport_tb.Items.Add(whItem);
            }

            if (wh_instructiontable != null && wh_instructiontable.Rows.Count > 0)
            {
                ComboBoxItem whItem = new ComboBoxItem
                {
                    Content = "Where-used_" + wh_instructioninput
                };
                rtExport_tb.Items.Add(whItem);
            }

            if (wh_alternatetable != null && wh_alternatetable.Rows.Count > 0)
            {
                ComboBoxItem whItem = new ComboBoxItem
                {
                    Content = "Alternate_" + wh_alterinput
                };
                rtExport_tb.Items.Add(whItem);
            }

            if (wh_materialtable != null && wh_materialtable.Rows.Count > 0)
            {
                ComboBoxItem whItem = new ComboBoxItem
                {
                    Content = "Material_" + wh_materialinput
                };
                rtExport_tb.Items.Add(whItem);
            }

            if (wh_depttable != null && wh_depttable.Rows.Count > 0)
            {
                ComboBoxItem whItem = new ComboBoxItem
                {
                    Content = "Dept_" + wh_deptinput
                };
                rtExport_tb.Items.Add(whItem);
            }

            if (ecn_table != null && ecn_table.Rows.Count > 0 && whECN.SelectedIndex == 0)
            {
                ComboBoxItem whItem = new ComboBoxItem
                {
                    Content = ecn_input
                };
                rtExport_tb.Items.Add(whItem);
            }
            else if (ecn_table != null && ecn_table.Rows.Count > 0 && whECN.SelectedIndex == 1)
            {
                ComboBoxItem whItem = new ComboBoxItem
                {
                    Content = "ECN_" + ecn_input
                };
                rtExport_tb.Items.Add(whItem);
            }
        }

        #region Button Click and Routing Function
        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            if (item_no.SelectedIndex == 0)
            {
                item_input = rtInput.Text.ToUpper();
            }
            else if (item_no.SelectedIndex == 1)
            {
                string[] tokens = rtInput.Text.Split(new char[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                HashSet<string> uniqueStrings = new HashSet<string>();

                foreach (string token in tokens)
                {
                    uniqueStrings.Add(token.Trim());
                }
                itemlist_input = uniqueStrings.OrderBy(s => s).ToList();
            }

            if (rtEcn_Input.Text != "")
            {
                ecn_input = rtEcn_Input.Text.ToUpper();
            }
            if (wh_Input.Text != "" && wh_select.SelectedIndex == 1)
            {
                wh_inputdoc = wh_Input.Text.ToUpper();
            }
            else if (wh_Input.Text != "" && wh_select.SelectedIndex == 2)
            {
                wh_instructioninput = wh_Input.Text.ToUpper();
            }
            else if (wh_Input.Text != "" && wh_select.SelectedIndex == 3)
            {
                wh_alterinput = wh_Input.Text.ToUpper();
            }
            else if (wh_Input.Text != "" && wh_select.SelectedIndex == 4)
            {
                wh_materialinput = wh_Input.Text.ToUpper();
            }
            else if (wh_Input.Text != "" && wh_select.SelectedIndex == 5)
            {
                wh_deptinput = wh_Input.Text.ToUpper();
            }

            //---------------------------------------------------------------
            //----ROUTING----------------------------------------------------
            //---------------------------------------------------------------

            if (rtInput.IsFocused || (rtECN.IsChecked == false && rtWhu.IsChecked == false))
            {
                if (rtInput.Text != "" && item_no.SelectedIndex == 0)
                {
                    rt_table.Clear();
                    bm_table.Clear();

                    if (rtRND.IsChecked == true)
                    {
                        item_type = "RND";
                        rtSTD.IsChecked = false;
                    }
                    else
                    {
                        item_type = "STD";
                        rtSTD.IsChecked = true;
                    }
                    Routing_Data(item_input, item_type);
                }
                else if (rtInput.Text != "" && item_no.SelectedIndex == 1)
                {
                    rt_table.Clear();
                    bm_table.Clear();

                    if (rtRND.IsChecked == true)
                    {
                        item_type = "RND";
                        rtSTD.IsChecked = false;
                    }
                    else
                    {
                        item_type = "STD";
                        rtSTD.IsChecked = true;
                    }
                    _ = RoutingList_DataAsync(itemlist_input, item_type);
                }
            }

            //---------------------------------------------------------------
            //----ECN--------------------------------------------------------
            //---------------------------------------------------------------

            if (rtEcn_Input.Text != "" || rtECN.IsChecked == true)
            {
                rtECN.IsChecked = true;
                ECN_Checked(sender, e);
                ECN_Data(ecn_input);
            }

            //---------------------------------------------------------------
            //----WHERE-USED-------------------------------------------------
            //---------------------------------------------------------------

            if (rtWhu.IsChecked == true)
            {
                if (rtWhu.IsChecked == true)
                {
                    if (item_input != "" && wh_Input.Text == "")
                    {
                        if (routing_gridview.ColumnDefinitions.Count < 7)
                        {
                            routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
                            routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
                        }

                        Grid.SetRow(whu_Grid_2, 0);
                        Grid.SetRowSpan(whu_Grid_2, 4);
                        lbWH_2.Margin = new Thickness(0, 64, 0, 0);
                        whuGrid_2.Margin = new Thickness(0, 90, 10, 10);

                        whu_Grid.Visibility = Visibility.Hidden;
                        whu_Grid_2.Visibility = Visibility.Visible;
                    }
                    else if (wh_Input.Text != "")
                    {
                        if (routing_gridview.ColumnDefinitions.Count < 7)
                        {
                            routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
                            routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
                        }

                        Grid.SetRow(whu_Grid_2, 0);
                        Grid.SetRowSpan(whu_Grid_2, 2);
                        lbWH_2.Margin = new Thickness(0, 64, 0, 0);
                        whuGrid_2.Margin = new Thickness(0, 90, 10, 0);

                        Grid.SetRow(whu_Grid, 2);
                        Grid.SetRowSpan(whu_Grid, 2);
                        lbWH.Margin = new Thickness(0, 0, 0, 0);
                        whuGrid.Margin = new Thickness(0, 30, 10, 10);

                        whu_Grid.Visibility = Visibility.Visible;
                        whu_Grid.Visibility = Visibility.Visible;
                    }
                    else if (item_input == "" && wh_Input.Text == "")
                    {
                        if (routing_gridview.ColumnDefinitions.Count < 7)
                        {
                            routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
                            routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
                        }

                        Grid.SetRow(whu_Grid, 0);
                        Grid.SetRowSpan(whu_Grid, 4);
                        lbWH.Margin = new Thickness(0, 64, 0, 0);
                        whuGrid.Margin = new Thickness(0, 90, 10, 10);

                        whu_Grid.Visibility = Visibility.Visible;
                        whu_Grid_2.Visibility = Visibility.Hidden;
                    }
                }
                if (wh_Input.Text != "" && wh_select.SelectedIndex == 1)
                {
                    Whereused_Data(wh_inputdoc, wh_doctable, whuGrid, lbWH);
                }
                else if (wh_Input.Text != "" && wh_select.SelectedIndex == 2)
                {
                    Whereused_Data(wh_instructioninput, wh_instructiontable, whuGrid, lbWH);
                }
                else if (wh_Input.Text != "" && wh_select.SelectedIndex == 3)
                {
                    Alter_function(wh_alterinput);
                }
                else if (wh_Input.Text != "" && wh_select.SelectedIndex == 4)
                {
                    Material_function(wh_materialinput);
                }
                else if (wh_Input.Text != "" && wh_select.SelectedIndex == 5)
                {
                    Dept_function(wh_deptinput);
                }
            }
        }

        private void RtExport_Click(object sender, RoutedEventArgs e)
        {
            if (rtExport_tb.SelectedItem != null)
            {
                string selectedTable = (rtExport_tb.SelectedItem as ComboBoxItem).Content.ToString();
                if (selectedTable == "Routing_" + item_input)
                {
                    ExcelTool.ExportExcelWithDialog(rt_table, item_input, "Routing_" + item_input);
                }
                else if (selectedTable == "Bom_" + item_input)
                {
                    ExcelTool.ExportExcelWithDialog(bm_table, item_input, "Bom_" + item_input);
                }
                else if (selectedTable == "Where-used_" + wh_inputmanual)
                {
                    ExcelTool.ExportExcelWithDialog(manual_table, wh_inputmanual, "Where-used_" + wh_inputmanual);
                }
                else if (selectedTable == "Where-used_" + wh_inputdoc)
                {
                    ExcelTool.ExportExcelWithDialog(wh_doctable, wh_inputdoc, "Where-used_" + wh_inputdoc);
                }
                else if (selectedTable == "Where-used_" + wh_instructioninput)
                {
                    ExcelTool.ExportExcelWithDialog(wh_instructiontable, wh_instructioninput, "Where-used_" + wh_instructioninput);
                }
                else if (selectedTable == "Material_" + wh_materialinput)
                {
                    ExcelTool.ExportExcelWithDialog(wh_materialtable, wh_materialinput, "Material_" + wh_materialinput);
                }
                else if (selectedTable == "Alternate_" + wh_alterinput)
                {
                    ExcelTool.ExportExcelWithDialog(wh_alternatetable, wh_alterinput, "Alternate_" + wh_alterinput);
                }
                else if (selectedTable == "Dept_" + wh_deptinput)
                {
                    ExcelTool.ExportExcelWithDialog(wh_depttable, wh_deptinput, "Dept_" + wh_deptinput);
                }
                else if (selectedTable == "ECN_" + ecn_input)
                {
                    ExcelTool.ExportExcelWithDialog(ecn_table, ecn_input, "ECN_" + ecn_input);
                }
                else if (selectedTable == ecn_input)
                {
                    ExcelTool.ExportExcelWithDialog(ecn_table, ecn_input, ecn_input);
                }
            }
        }

        private void Routing_Click(object sender, MouseButtonEventArgs e)
        {
            DataRowView dataRow = (DataRowView)routingGrid.SelectedItem;
            int index = routingGrid.CurrentCell.Column.DisplayIndex;
            string cellValue = dataRow.Row.ItemArray[index].ToString();

            if (cellValue == " " || cellValue == "" || cellValue == null)
            {
                MessageBox.Show("Cell value is null.");
            }
            else
            {
                if (index == 4)
                {
                    wh_inputmanual = cellValue;
                    if (rtWhu.IsChecked == false)
                    {
                        Process.Start(PathReader.EDM_link + wh_inputmanual);
                    }
                    else if (rtWhu.IsChecked == true)
                    {
                        Whereused_Data(wh_inputmanual, manual_table, whuGrid_2, lbWH_2);
                    }
                }
                if (index == 8)
                {
                    ecn_input = cellValue;
                    rtECN.IsChecked = true;
                    ECN_Checked(sender, e);
                    ECN_Data(ecn_input);
                }
            }
        }

        public void Routing_Data(string input, string type)
        {
            try
            {
                mv_cnt = new OdbcConnection
                {
                    ConnectionString = PathReader.ODSSG_server
                };
                mv_cnt.Open();
                Console.WriteLine("Connect successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Connect error: {ex.Message}");
            }

            try
            {
                if (input == "")
                {
                    MessageBox.Show("Please enter the value you want to find.");
                }
                else
                {
                    OdbcCommand rt_cmd = new OdbcCommand
                    {
                        Connection = mv_cnt,
                        CommandType = CommandType.Text
                    };
                    if (rtAll.IsChecked == false)
                    {
                        rt_cmd.CommandText = CommandText_1;
                        rt_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();
                        rt_cmd.Parameters.Add("@type", OdbcType.Char).Value = type;
                    }
                    else if (rtAll.IsChecked == true)
                    {
                        rt_cmd.CommandText = CommandText_2;
                        rt_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();
                        rt_cmd.Parameters.Add("@type", OdbcType.Char).Value = type;
                    }

                    OdbcDataAdapter rt_data = new OdbcDataAdapter
                    {
                        SelectCommand = rt_cmd
                    };
                    rt_data.Fill(rt_table);
                    item_No = rt_table.Rows[0]["POPRNO"].ToString();

                    DataTable tt_dt = new DataTable();
                    OdbcCommand tt_cmd = new OdbcCommand
                    {
                        Connection = mv_cnt,
                        CommandType = CommandType.Text
                    };
                    tt_cmd.CommandText = "Select MMFUDS from PFODS.MITMAS_143 where MMITNO = '" + input.Trim() + "'";
                    OdbcDataAdapter getDataData = new OdbcDataAdapter();
                    getDataData.SelectCommand = tt_cmd;
                    getDataData.Fill(tt_dt);
                    string title = "";
                    foreach (DataRow row in tt_dt.Rows)
                    {
                        title = row["MMFUDS"].ToString();
                    }

                    lbRT.Content = "ROUTING: " + input;
                    Title = "Movex Infor: " + item_No + "--" + title;
                    DateTimeFormat.DatetimeFormat(rt_table, routingGrid, "A");

                    OdbcCommand bm_cmd = new OdbcCommand
                    {
                        Connection = mv_cnt,
                        CommandType = CommandType.Text
                    };
                    if (rtAll.IsChecked == false)
                    {
                        bm_cmd.CommandText = CommandText_3;
                        bm_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();
                        bm_cmd.Parameters.Add("@type", OdbcType.Char).Value = type;
                    }
                    else if (rtAll.IsChecked == true)
                    {
                        bm_cmd.CommandText = CommandText_4;
                        bm_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();
                        bm_cmd.Parameters.Add("@type", OdbcType.Char).Value = type;
                    }

                    OdbcDataAdapter bm_data = new OdbcDataAdapter
                    {
                        SelectCommand = bm_cmd
                    };
                    bm_data.Fill(bm_table);
                    lbBM.Content = "BOM: " + input;
                    DateTimeFormat.DatetimeFormat(bm_table, bomGrid, "A");
                }
                mv_cnt.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Item number is not exit.");
            }
        }

        public async Task RoutingList_DataAsync(List<string> inputs, string type)
        {
            try
            {
                mv_cnt = new OdbcConnection
                {
                    ConnectionString = PathReader.ODSSG_server
                };
                mv_cnt.Open();
                Console.WriteLine("Connect successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Connect error: {ex.Message}");
            }

            ProgressBox prog = new ProgressBox();
            prog.Show();
            UpdateProgressDelegate updateProgress = new UpdateProgressDelegate(prog.UpdateProgress);
            int count = 0;
            int percent;

            try
            {
                if (inputs.Count == 0)
                {
                    MessageBox.Show("Please enter the value you want to find.");
                }
                else
                {
                    foreach (string input in inputs)
                    {
                        count++;
                        percent = count * 100 / inputs.Count();

                        updateProgress.Invoke(Convert.ToInt32(percent), percent.ToString());
                        await Task.Delay(10);

                        OdbcCommand rt_cmd = new OdbcCommand
                        {
                            Connection = mv_cnt,
                            CommandType = CommandType.Text
                        };
                        if (rtAll.IsChecked == false)
                        {
                            rt_cmd.CommandText = CommandText_1;
                            rt_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();
                            rt_cmd.Parameters.Add("@type", OdbcType.Char).Value = type;
                        }
                        else if (rtAll.IsChecked == true)
                        {
                            rt_cmd.CommandText = CommandText_2;
                            rt_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();
                            rt_cmd.Parameters.Add("@type", OdbcType.Char).Value = type;
                        }

                        OdbcDataAdapter rt_data = new OdbcDataAdapter
                        {
                            SelectCommand = rt_cmd
                        };
                        rt_data.Fill(rt_table);

                        OdbcCommand bm_cmd = new OdbcCommand
                        {
                            Connection = mv_cnt,
                            CommandType = CommandType.Text
                        };
                        if (rtAll.IsChecked == false)
                        {
                            bm_cmd.CommandText = CommandText_3;
                            bm_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();
                            bm_cmd.Parameters.Add("@type", OdbcType.Char).Value = type;
                        }
                        else if (rtAll.IsChecked == true)
                        {
                            bm_cmd.CommandText = CommandText_4;
                            bm_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();
                            bm_cmd.Parameters.Add("@type", OdbcType.Char).Value = type;
                        }

                        OdbcDataAdapter bm_data = new OdbcDataAdapter
                        {
                            SelectCommand = bm_cmd
                        };
                        bm_data.Fill(bm_table);
                    }
                    prog.Close();
                    lbRT.Content = "ROUTING: list items";
                    lbBM.Content = "BOM: list items";

                    string result = string.Join(", ", inputs);
                    Title = "Movex Infor: " + result;
                    item_input = "list items";

                    DateTimeFormat.DatetimeFormat(rt_table, routingGrid, "A");
                    DateTimeFormat.DatetimeFormat(bm_table, bomGrid, "A");
                }
                mv_cnt.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Item number is not exit.");
                prog.Close();
            }
        }
        #endregion

        #region Where-used Function
        private void Wh_SeclectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (rtWhu.IsChecked == true)
            {
                if (wh_select.SelectedIndex == 0)
                {
                    if (routing_gridview.ColumnDefinitions.Count < 10)
                    {
                        routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
                        routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
                        routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
                    }

                    Grid.SetRow(whu_Grid_2, 0);
                    Grid.SetRowSpan(whu_Grid_2, 4);
                    lbWH_2.Margin = new Thickness(0, 64, 0, 0);
                    whuGrid_2.Margin = new Thickness(0, 90, 10, 10);

                    whu_Grid.Visibility = Visibility.Hidden;
                    whu_Grid_2.Visibility = Visibility.Visible;
                    cb_2.Visibility = Visibility.Hidden;

                    wh_Input.IsEnabled = false;
                }
                else if (wh_select.SelectedIndex == 1)
                {
                    wh_Input.IsEnabled = true;
                    wh_Input.Focus();
                    wh_Input.SelectAll();
                    cb_2.Visibility = Visibility.Hidden;
                }
                else if (wh_select.SelectedIndex == 2)
                {
                    wh_Input.IsEnabled = true;
                    wh_Input.Focus();
                    wh_Input.SelectAll();
                    cb_2.Visibility = Visibility.Hidden;
                }
                else if (wh_select.SelectedIndex == 3)
                {
                    wh_Input.IsEnabled = true;
                    cb_2.Margin = new Thickness(0, 64, 10, 0);
                    cb_2.Visibility = Visibility.Visible;
                }
                else if (wh_select.SelectedIndex == 4)
                {
                    wh_Input.IsEnabled = true;
                    wh_Input.Focus();
                    wh_Input.SelectAll();
                    cb_2.Visibility = Visibility.Hidden;
                }
                else if (wh_select.SelectedIndex == 5)
                {
                    wh_Input.IsEnabled = true;
                    wh_Input.Text = "VP-";
                    wh_Input.Focus();
                    wh_Input.SelectAll();
                    cb_2.Visibility = Visibility.Hidden;
                }
            }
        }

        public async void Whereused_Data(string input, DataTable dt, DataGrid dtgr, Label label)
        {
            mv_cnt = new OdbcConnection
            {
                ConnectionString = PathReader.ODSSG_server
            };
            mv_cnt.Open();

            OdbcCommand rt_cmd = new OdbcCommand
            {
                Connection = mv_cnt,
                CommandType = CommandType.Text
            };
            if (wh_select.SelectedIndex != 0)
            {
                if (wh_select.SelectedIndex == 1)
                {
                    ProcessWhu(rt_cmd, input, dt, label, dtgr);
                }
                else if (wh_select.SelectedIndex == 2)
                {
                    await ProcessWhu58(rt_cmd, input, dt, label, dtgr);
                }
            }
            else
            {
                ProcessWhu(rt_cmd, input, dt, label, dtgr);
            }

            mv_cnt.Close();
        }

        private void ProcessWhu(OdbcCommand rt_cmd, string input, DataTable dt, Label label, DataGrid dtgr)
        {
            rt_cmd.CommandText = CommandText_5;
            rt_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();

            OdbcDataAdapter rt_data = new OdbcDataAdapter
            {
                SelectCommand = rt_cmd
            };
            dt.Clear();
            rt_data.Fill(dt);

            label.Content = "WHERE-USED: " + input;
            DateTimeFormat.DatetimeFormat(dt, dtgr, "A");
        }

        private async Task ProcessWhu58(OdbcCommand rt_cmd, string input, DataTable dt, Label label, DataGrid dtgr)
        {
            ProgressBox prog = new ProgressBox();
            prog.Show();

            try
            {
                UpdateProgressDelegate updateProgress = new UpdateProgressDelegate(prog.UpdateProgress);
                updateProgress.Invoke(10, "10");
                await Task.Delay(10);

                rt_cmd.CommandText = CommandText_6;
                rt_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();

                OdbcDataAdapter rt_data = new OdbcDataAdapter
                {
                    SelectCommand = rt_cmd
                };
                dt.Clear();
                rt_data.Fill(dt);
                updateProgress.Invoke(30, "30");
                await Task.Delay(10);

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("Cannot find any relatived items, please try again with another index");
                    prog.Close();
                    return;
                }

                DataTable uniqueDt = dt.DefaultView.ToTable(true, "POPRNO");
                dt.Clear();
                foreach (DataRow row in uniqueDt.Rows)
                {
                    DataRow newRow = dt.NewRow();
                    newRow["POPRNO"] = row["POPRNO"];
                    dt.Rows.Add(newRow);
                }
                List<DataTable> listOfTables = new List<DataTable>();
                updateProgress.Invoke(70, "70");
                await Task.Delay(10);

                foreach (DataRow row in dt.Rows)
                {
                    DataTable tempTable = await ProcessWhu58Detail(row["POPRNO"].ToString());
                    listOfTables.Add(tempTable);
                }
                updateProgress.Invoke(90, "90");
                await Task.Delay(10);
                DataTable rs_table = ProcessWhu58Result(listOfTables, input);
                dt = rs_table;
                updateProgress.Invoke(100, "100");
                await Task.Delay(10);
                label.Content = "INSTRUCTION: " + input;
                DateTimeFormat.DatetimeFormat(dt, dtgr, "A");
                prog.Close();
            }
            catch
            {
                prog.Close();
            }
        }

        private async Task<DataTable> ProcessWhu58Detail(string poprno)
        {
            OdbcCommand rt_cmd = new OdbcCommand
            {
                Connection = mv_cnt,
                CommandType = CommandType.Text
            };

            DataTable tempTable = new DataTable();
            rt_cmd.CommandText = CommandText_7;
            rt_cmd.Parameters.Add("@poprno", OdbcType.Char).Value = poprno;

            OdbcDataAdapter rt_dt = new OdbcDataAdapter
            {
                SelectCommand = rt_cmd
            };

            await Task.Run(() => rt_dt.Fill(tempTable));

            return tempTable;
        }

        private DataTable ProcessWhu58Result(List<DataTable> listOfTables, string input)
        {
            DataTable rs_table = null;
            foreach (DataTable dt_x in listOfTables)
            {
                bool flag = false;

                for (int i = dt_x.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow currentRow = dt_x.Rows[i];

                    if (currentRow["PODOID"].ToString() == input.Trim())
                    {
                        flag = true;
                    }

                    if (flag && Convert.ToInt32(currentRow["POPITI"]) != 0)
                    {
                        if (rs_table == null)
                        {
                            rs_table = dt_x.Clone();
                        }

                        rs_table.Rows.Add(currentRow.ItemArray);
                        flag = false;
                    }
                }
            }
            return rs_table;
        }
        #endregion

        public void ECN_Data(string input)
        {
            ecn_table.Clear();
            try
            {
                mv_cnt = new OdbcConnection
                {
                    ConnectionString = PathReader.ODSSG_server
                };
                mv_cnt.Open();

                if (whECN.SelectedIndex == 0)
                {
                    OdbcCommand rt_cmd = new OdbcCommand
                    {
                        Connection = mv_cnt,
                        CommandType = CommandType.Text,
                        CommandText = CommandText_8
                    };
                    rt_cmd.Parameters.Add("@poprno", OdbcType.Char).Value = input.Trim();

                    OdbcDataAdapter rt_data = new OdbcDataAdapter
                    {
                        SelectCommand = rt_cmd
                    };
                    rt_data.Fill(ecn_table);

                    lbECN.Content = "ECN to Items: " + input;
                }
                else if (whECN.SelectedIndex == 1)
                {
                    OdbcCommand rt_cmd = new OdbcCommand
                    {
                        Connection = mv_cnt,
                        CommandType = CommandType.Text,
                        CommandText = CommandText_8A
                    };
                    rt_cmd.Parameters.Add("@EBITNO", OdbcType.Char).Value = input.Trim();
                    rt_cmd.Parameters.Add("@EBITNN", OdbcType.Char).Value = input.Trim();

                    OdbcDataAdapter rt_data = new OdbcDataAdapter
                    {
                        SelectCommand = rt_cmd
                    };
                    rt_data.Fill(ecn_table);
                    lbECN.Content = "Item to ECNs: " + input;
                }

                DateTimeFormat.DatetimeFormat(ecn_table, ecnGrid, "A");
                rtEcn_Input.Text = "";
                mv_cnt.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Item number is not exit." + e);
            }
        }

        public void Dept_function(string input)
        {
            try
            {
                mv_cnt = new OdbcConnection
                {
                    ConnectionString = PathReader.ODSSG_server
                };
                mv_cnt.Open();

                OdbcCommand rt_cmd = new OdbcCommand
                {
                    Connection = mv_cnt,
                    CommandType = CommandType.Text
                };

                wh_depttable.Clear();
                rt_cmd.CommandText = CommandText_12;
                rt_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();

                OdbcDataAdapter rt_dt = new OdbcDataAdapter
                {
                    SelectCommand = rt_cmd
                };
                rt_dt.Fill(wh_depttable);

                lbWH_2.Content = "DEPT: " + input;
                DateTimeFormat.DatetimeFormat(wh_depttable, whuGrid_2, "A");
                mv_cnt.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        #region CheckBox
        private void WhereUse_Checked(object sender, RoutedEventArgs e)
        {
            if (routing_gridview.ColumnDefinitions.Count < 10)
            {
                routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
                routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
                routing_gridview.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(1, GridUnitType.Star) });
            }
            wh_select.IsEnabled = true;

            Grid.SetRow(whu_Grid_2, 0);
            Grid.SetRowSpan(whu_Grid_2, 4);
            lbWH_2.Margin = new Thickness(0, 64, 0, 0);
            whuGrid_2.Margin = new Thickness(0, 90, 10, 10);

            whu_Grid.Visibility = Visibility.Hidden;
            whu_Grid_2.Visibility = Visibility.Visible;
        }

        private void WhereUse_UnChecked(object sender, RoutedEventArgs e)
        {
            if (routing_gridview.ColumnDefinitions.Count > 7)
            {
                routing_gridview.ColumnDefinitions.RemoveAt(routing_gridview.ColumnDefinitions.Count - 1);
                routing_gridview.ColumnDefinitions.RemoveAt(routing_gridview.ColumnDefinitions.Count - 1);
                routing_gridview.ColumnDefinitions.RemoveAt(routing_gridview.ColumnDefinitions.Count - 1);
            }
            whu_Grid.Visibility = Visibility.Hidden;
            whu_Grid_2.Visibility = Visibility.Hidden;

            wh_select.SelectedIndex = 0;
            wh_select.IsEnabled = false;
            wh_Input.Text = "";
        }

        private void ECN_Checked(object sender, RoutedEventArgs e)
        {
            expand_ecn = true;
            whECN.IsEnabled = true;
            Grid.SetRowSpan(bom_Grid, 2);
            bomGrid.Margin = new Thickness(0, 90, 10, 0);
            if (whECN.SelectedIndex == 0)
            {
                rtEcn_Input.Text = "ECN-";
            }
            else
            {
                rtEcn_Input.Text = "";
            }
            rtEcn_Input.IsEnabled = true;
            rtEcn_Input.Focus();
            rtEcn_Input.CaretIndex = rtEcn_Input.Text.Length;
            ecn_Grid.Visibility = Visibility.Visible;
        }

        private void ECN_Unchecked(object sender, RoutedEventArgs e)
        {
            expand_ecn = false;
            whECN.IsEnabled = false;
            Grid.SetRowSpan(bom_Grid, 4);
            bomGrid.Margin = new Thickness(0, 90, 10, 10);
            rtEcn_Input.Text = "";
            rtEcn_Input.IsEnabled = false;
            ecn_Grid.Visibility = Visibility.Hidden;
        }

        private void whECN_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (whECN.SelectedIndex == 1)
            {
                rtEcn_Input.Text = "";
                rtEcn_Input.Focus();
            }
            else if (rtEcn_Input != null)
            {
                rtEcn_Input.Text = "ECN-";
                rtEcn_Input.Select(rtEcn_Input.Text.Length, 0);
                rtEcn_Input.Focus();
            }
        }
        #endregion

        #region Alter, Dept & Material Function
        private void Alter_Click(object sender, MouseButtonEventArgs e)
        {
            DataRowView dataRow = (DataRowView)bomGrid.SelectedItem;
            int index = bomGrid.CurrentCell.Column.DisplayIndex;
            string cellValue = dataRow.Row.ItemArray[index].ToString();

            if (cellValue == " " || cellValue == "" || cellValue == null)
            {
                MessageBox.Show("Cell value is null.");
            }
            else
            {
                if (index == 4 && wh_select.SelectedIndex == 3)
                {
                    wh_alterinput = cellValue;
                    alter = dataRow.Row.ItemArray[index - 1].ToString();
                    Alter_function(wh_alterinput);
                }
                else if (index == 4 && wh_select.SelectedIndex == 4)
                {
                    wh_materialinput = cellValue;
                    Material_function(wh_materialinput);
                }
                else
                {
                    MessageBox.Show("No function check.");
                }
            }
        }

        public async void Alter_function(string input)
        {
            ProgressBox prog = new ProgressBox();
            prog.Show();

            UpdateProgressDelegate updateProgress = new UpdateProgressDelegate(prog.UpdateProgress);

            try
            {
                mv_cnt = new OdbcConnection
                {
                    ConnectionString = PathReader.ODSSG_server
                };
                mv_cnt.Open();

                OdbcCommand rt_cmd = new OdbcCommand
                {
                    Connection = mv_cnt,
                    CommandType = CommandType.Text
                };

                DataTable itemTable = new DataTable();

                rt_cmd.CommandText = CommandText_9;
                rt_cmd.Parameters.Add("@poprno", OdbcType.Char).Value = input.Trim();

                OdbcDataAdapter rt_dt = new OdbcDataAdapter
                {
                    SelectCommand = rt_cmd
                };
                rt_dt.Fill(itemTable);

                updateProgress.Invoke(20, "20");
                await Task.Delay(10);

                DataTable temp_tb = new DataTable();
                double x = 50.0 / itemTable.Rows.Count;
                double y = 20.0;
                double z = 20.0;

                foreach (DataRow n in itemTable.Rows)
                {
                    string m = n[0].ToString();

                    if (x + y - z >= 1.0)
                    {
                        z = x + y;
                        int intValue = (int)Math.Floor(z);

                        updateProgress.Invoke(intValue, intValue.ToString());
                        await Task.Delay(10);
                    }

                    y += x;

                    OdbcCommand cmd = new OdbcCommand
                    {
                        Connection = mv_cnt,
                        CommandType = CommandType.Text,
                        CommandText = CommandText_10
                    };
                    cmd.Parameters.Add("@poprno", OdbcType.Char).Value = m;

                    OdbcDataAdapter dt = new OdbcDataAdapter
                    {
                        SelectCommand = cmd
                    };
                    dt.Fill(temp_tb);
                }

                DataTable alterTable = temp_tb.Clone();
                for (int i = 0; i < temp_tb.Rows.Count; i++)
                {
                    DataRow row = temp_tb.Rows[i];
                    if (row["PMMTNO"].ToString() == input)
                    {
                        string A = row["PMPRNO"].ToString();
                        string B = row["PMDWPO"].ToString();
                        string[] B1 = B.Split(' ');
                        string C = B1[0];
                        string D = B1.Length > 1 ? B1[1] : "-";

                        string set = A + C;
                        bool check = false;

                        foreach (DataRow n in temp_tb.Rows)
                        {
                            string E = n["PMPRNO"].ToString();
                            string F = n["PMDWPO"].ToString();
                            string[] F1 = F.Split(' ');

                            string G = F1[0];
                            string H = F1.Length > 1 ? F1[1] : "-";

                            string compare = E + G;

                            if (compare == set && H != D)
                            {
                                alterTable.ImportRow(n);
                                check = true;
                            }
                        }

                        if (check)
                        {
                            alterTable.ImportRow(row);
                        }
                    }
                }

                updateProgress.Invoke(80, "80");
                await Task.Delay(10);

                RemoveDuplicateRows(alterTable, "PMPRNO");
                RemoveDuplicateValue(alterTable, "PMMTNO", input);

                updateProgress.Invoke(90, "90");
                await Task.Delay(10);

                lbWH_2.Content = "ALTERNATE: " + input;
                cb_2.Visibility = Visibility.Visible;
                DateTimeFormat.DatetimeFormat(alterTable, whuGrid_2, "A");
                mv_cnt.Close();

                updateProgress.Invoke(100, "100");
                await Task.Delay(10);

                prog.Close();
            }
            catch (Exception ex)
            {
                prog.Close();
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        /*public void Material_function(string input)
        {
            mv_cnt = new OdbcConnection
            {
                ConnectionString = PathReader.ODSSG_server
            };
            mv_cnt.Open();

            OdbcDataAdapter data = Material_data(input, mv_cnt);
            wh_materialtable.Clear();

            DataTable temp = new DataTable();
            if (!temp.Columns.Contains("Level"))
            {
                temp.Columns.Add("Level", typeof(int));
            }
            data.Fill(temp);
            wh_materialtable = temp.Clone();

            foreach (DataRow row in temp.Rows)
            {
                row["Level"] = 1;
                wh_materialtable.ImportRow(row);
                DataTable temp_n = new DataTable();
                OdbcDataAdapter data_n = Material_data(row[1].ToString(), mv_cnt);
                temp_n = temp.Clone();
                data_n.Fill(temp_n);

                if (temp_n.Rows.Count > 0)
                {
                    foreach (DataRow row_1 in temp_n.Rows)
                    {
                        row_1["Level"] = 2;
                        wh_materialtable.ImportRow(row_1);
                        DataTable temp_n1 = new DataTable();
                        OdbcDataAdapter data_n1 = Material_data(row_1[1].ToString(), mv_cnt);
                        temp_n1 = temp.Clone();
                        data_n1.Fill(temp_n1);
                        if (temp_n1.Rows.Count > 0)
                        {
                            foreach (DataRow row_2 in temp_n1.Rows)
                            {
                                row_2["Level"] = 3;
                                wh_materialtable.ImportRow(row_2);
                            }
                        }
                    }
                }
            }

            temp.Clear();
            lbWH_2.Content = "Material No." + input;
            DateTimeFormat.DatetimeFormat(wh_materialtable, whuGrid_2, "A");
            mv_cnt.Close();
        }*/

        public void Material_function(string input)
        {
            mv_cnt = new OdbcConnection
            {
                ConnectionString = PathReader.ODSSG_server
            };
            mv_cnt.Open();

            OdbcDataAdapter data = Material_data(input, mv_cnt);
            wh_materialtable.Clear();

            DataTable temp = new DataTable();
            if (!temp.Columns.Contains("Level"))
            {
                temp.Columns.Add("Level", typeof(int));
            }
            data.Fill(temp);
            wh_materialtable = temp.Clone();

            Func<DataTable, bool> IsTableEmpty = (table) => table.Rows.Count == 0;
            Action<DataTable, int> ProcessLevel = null;
            ProcessLevel = (table, level) =>
            {
                foreach (DataRow row in table.Rows)
                {
                    row["Level"] = level;
                    wh_materialtable.ImportRow(row);
                    DataTable temp_n = new DataTable();
                    OdbcDataAdapter data_n = Material_data(row[1].ToString(), mv_cnt);
                    temp_n = temp.Clone();
                    data_n.Fill(temp_n);

                    if (!IsTableEmpty(temp_n))
                    {
                        ProcessLevel(temp_n, level + 1);
                    }
                }
            };

            if (!IsTableEmpty(temp))
            {
                ProcessLevel(temp, 1);
            }

            temp.Clear();
            lbWH_2.Content = "Material No." + input;
            DateTimeFormat.DatetimeFormat(wh_materialtable, whuGrid_2, "A");
            mv_cnt.Close();
        }

        public OdbcDataAdapter Material_data(string input, OdbcConnection cnt)
        {
            OdbcCommand cmd = new OdbcCommand
            {
                Connection = mv_cnt,
                CommandType = CommandType.Text
            };

            cmd.CommandText = CommandText_11;
            cmd.Parameters.Add("@item", OdbcType.Char).Value = input.Trim();

            OdbcDataAdapter data = new OdbcDataAdapter();
            data.SelectCommand = cmd;
            return data;
        }

        public void RemoveDuplicateValue(DataTable table, string columnName, string input)
        {
            HashSet<object> uniqueValues = new HashSet<object>();
            uniqueValues.Clear();
            cb_2.Items.Clear();

            foreach (DataRow row in table.Rows)
            {
                object value = row[columnName];
                if (value.ToString() != input)
                {
                    if (uniqueValues.Add(value))
                    {
                        cb_2.Items.Add(value);
                    }
                }
            }
            cb_2.SelectedIndex = 0;
        }

        static void RemoveDuplicateRows(DataTable table, string columnName)
        {
            DataView view = new DataView(table)
            {
                Sort = columnName
            };

            DataTable uniqueTable = view.ToTable(true, table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray());
            table.Rows.Clear();

            foreach (DataRow row in uniqueTable.Rows)
            {
                table.ImportRow(row);
            }
            wh_alternatetable = table;
        }
        #endregion
    }
}