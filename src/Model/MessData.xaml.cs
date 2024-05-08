using MnS.lib;
using System.Collections.Generic;
using System.Data;
using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using System.Data.Odbc;
using System.Windows.Controls;
using System.Windows.Input;
using System.Linq;
using OxyPlot.Series;
using System.Diagnostics;
using System.Threading.Tasks;

namespace MnS
{
    public partial class MessData : Window
    {
        #region Variable
        public static class Globalvariant
        {
            public static string database, user, password, service, SMT, PA, FA, Highvolt = null;
        }
        public static DataTable StepTable { get; set; }
        public static List<string> SelectedFiles { get; set; }

        public delegate void UpdateProgressDelegate(int value, string prog);
        public static List<string> Prog = new List<string>();
        public static List<string> PSL = new List<string>();
        public static List<DataTable> boxplot_dt = new List<DataTable>();

        public static List<Data_Infor> dataInfoList = new List<Data_Infor>();
        public static List<DataTable> dataTableList = new List<DataTable>();
        public static List<string> filePathList = new List<string>();
        public static DataTable mergeTable = new DataTable();

        public string column_Name;
        public string Min_Value;
        public string Max_Value;
        public string Unit;
        public string Sync;
        #endregion

        #region Item search CheckBox
        private void Mescheck(CheckBox check)
        {
            List<CheckBox> checklist = new List<CheckBox>
            {
                ms_HV,
                ms_FA,
                ms_PA,
                ms_SMT
            };
            CheckAndUncheck(check, checklist);
        }

        public static void CheckAndUncheck(CheckBox chosebox, List<CheckBox> hidenboxes)
        {
            if (chosebox == null || hidenboxes == null)
            {
                return;
            }

            hidenboxes.Remove(chosebox);
            foreach (CheckBox check in hidenboxes)
            {
                if (check != null && check.IsChecked == true)
                {
                    check.IsChecked = false;
                }
            }
            hidenboxes.Add(chosebox);
        }

        private void HV_Checked(object sender, RoutedEventArgs e)
        {
            Mescheck(ms_HV);
        }

        private void FA_Checked(object sender, RoutedEventArgs e)
        {
            Mescheck(ms_FA);
        }

        private void PA_Checked(object sender, RoutedEventArgs e)
        {
            Mescheck(ms_PA);
        }

        private void SMT_Checked(object sender, RoutedEventArgs e)
        {
            Mescheck(ms_SMT);
        }

        private string GetFolder()
        {
            string messFolder = "";
            if (ms_SMT.IsChecked == true)
            {
                messFolder = Globalvariant.SMT;
            }
            if (ms_FA.IsChecked == true)
            {
                messFolder = Globalvariant.FA;
            }
            if (ms_PA.IsChecked == true)
            {
                messFolder = Globalvariant.PA;
            }
            if (ms_HV.IsChecked == true)
            {
                messFolder = Globalvariant.Highvolt;
            }
            return messFolder;
        }
        #endregion

        #region Select CheckBox
        private void Selectcheck(CheckBox check)
        {
            List<CheckBox> checklist = new List<CheckBox>
            {
                msSingle, msDoub, msMulti
            };
            CheckAndUncheck(check, checklist);
        }

        private void SingleChecked(object sender, RoutedEventArgs e)
        {
            Selectcheck(msSingle);
            mes_get.IsEnabled = true;
            msPath.IsEnabled = true;
            mes_DouGrid1.IsEnabled = false;
            ms_PSL.IsEnabled = false;
            ms_TestProg.IsEnabled = false;
            ms_Step.IsEnabled = true;
            mes_List.IsEnabled = false;
            line_but.IsEnabled = true;
            box_but.IsEnabled = true;
            ClearData(sender, e);
        }

        private void MultiChecked(object sender, RoutedEventArgs e)
        {
            Selectcheck(msMulti);
            mes_get.IsEnabled = true;
            msPath.IsEnabled = false;
            mes_DouGrid1.IsEnabled = false;
            ms_PSL.IsEnabled = false;
            ms_TestProg.IsEnabled = false;
            ms_Step.IsEnabled = false;
            mes_List.IsEnabled = true;
            line_but.IsEnabled = false;
            box_but.IsEnabled = false;
            ClearData(sender, e);
        }

        private void DoubleChecked(object sender, RoutedEventArgs e)
        {
            Selectcheck(msDoub);
            mes_get.IsEnabled = false;
            msPath.IsEnabled = false;
            mes_DouGrid1.IsEnabled = true;
            ms_PSL.IsEnabled = true;
            ms_TestProg.IsEnabled = true;
            ms_Step.IsEnabled = true;
            mes_List.IsEnabled = true;
            line_but.IsEnabled = true;
            box_but.IsEnabled = true;
            ClearData(sender, e);
        }
        #endregion

        public MessData()
        {
            UserLogTool.UserData("Using Mess Data function");
            InitializeComponent();
            ms_Step.SelectionChanged += MsStep_SelectionChanged;
            ms_TestProg.SelectionChanged += ProgSelect_SelectionChanged;
        }

        private void ClearData(object sender, RoutedEventArgs e)
        {
            try
            {
                DatTool.DeleteAllFiles(PathReader.boxplot_foldercp);
                DatTool.DeleteAllFiles(PathReader.boxplot_foldermes);
                mergeTable.Rows.Clear();
                mergeTable.Columns.Clear();
                mes_List.Items.Clear();
                dataInfoList.Clear();
                dataTableList.Clear();
                filePathList.Clear();
                Sync = "";
                msPath.Text = "";
                ms_Section.Text = "";
                ms_PSL.ItemsSource = null;
                ms_PSL.Items.Clear();
                ms_TestProg.ItemsSource = null;
                ms_TestProg.Items.Clear();
                ms_Step.Items.Clear();
                Prog.Clear();
                PSL.Clear();
                boxplot_dt.Clear();
                ms_minv.Text = "-";
                ms_maxv.Text = "-";
                ms_unitv.Text = "-";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void OpenFiles(object sender, RoutedEventArgs e)
        {
            try
            {
                if (msSingle.IsChecked == false && msDoub.IsChecked == false && msMulti.IsChecked == false)
                {
                    MessageBox.Show("Please select a mode before searching.");
                }
                else
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog
                    {
                        Filter = "DAT Files|*.DAT|All Files|*.*",
                        InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    };

                    if (msMulti.IsChecked == true)
                    {
                        openFileDialog.Multiselect = true;
                    }
                    bool? result = openFileDialog.ShowDialog();
                    if (result == true)
                    {
                        if (SelectedFiles != null)
                        {
                            SelectedFiles.Clear();
                        }
                        ClearData(sender, e);
                        SelectedFiles = new List<string>();
                        SelectedFiles.AddRange(openFileDialog.FileNames);
                        DatTool.CopyFilesToFolder(SelectedFiles, PathReader.boxplot_foldermes);
                        foreach (string name in SelectedFiles)
                        {
                            string newname = Path.GetFileNameWithoutExtension(name);
                            DataTable table = new DataTable
                            {
                                TableName = newname
                            };
                            if (SelectedFiles.Count == 1)
                            {
                                msPath.Text = newname;
                            }
                            else
                            {
                                mes_List.Items.Add(newname);
                            }
                            DatTool.ReadFile(name, table);
                        }
                        if (DataProcess())
                        {
                            Sync = "synchronize";
                            ms_PSL.IsEnabled = true;
                            ms_TestProg.IsEnabled = true;
                            ms_Step.IsEnabled = true;
                            ms_BoxAll.IsEnabled = true;
                            SelectTestProg();
                            SelectTestStep();
                            line_but.IsEnabled = true;
                            box_but.IsEnabled = true;
                            MessageBox.Show("Synchronized data, you can export to an Excel file or view in a value chart.");
                        }
                        else
                        {
                            Sync = "asynchronous";
                            ms_PSL.IsEnabled = false;
                            ms_TestProg.IsEnabled = false;
                            ms_Step.IsEnabled = false;
                            ms_BoxAll.IsEnabled = false;
                            line_but.IsEnabled = false;
                            box_but.IsEnabled = false;
                            MessageBox.Show("Asynchronous data, you can only export to an Excel file and cannot view in a value chart.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private bool DataProcess()
        {
            if (SelectedFiles.Count > 1)
            {
                string check1 = dataInfoList[0].TestNumber;
                string check2 = dataInfoList[0].PartNumber;

                for (int i = 0; i < dataTableList.Count; i++)
                {
                    if (dataInfoList[i].TestNumber != check1 || dataInfoList[i].PartNumber != check2)
                    {
                        return false;
                    }
                    check1 = dataInfoList[i].TestNumber;
                    check2 = dataInfoList[i].PartNumber;
                }
            }
            return true;
        }

        private void ExportExcel(object sender, RoutedEventArgs e)
        {
            mergeTable.Columns.Clear();
            mergeTable.Rows.Clear();
            ConvertData();
        }

        private void ConvertData()
        {
            try
            {
                if (Sync == "synchronize")
                {
                    for (int i = 0; i < dataTableList.Count; i++)
                    {
                        DataTable dt = dataTableList[i];
                        Data_Infor dt_i = dataInfoList[i];

                        int moNoColumnIndex = dt.Columns.IndexOf("MO_no.");
                        if (moNoColumnIndex == -1)
                        {
                            DataColumn newColumn = new DataColumn("MO_no.", typeof(string));
                            dt.Columns.Add(newColumn);
                            dt.Columns["MO_no."].SetOrdinal(0);
                        }

                        foreach (DataRow row in dt.Rows)
                        {
                            row["MO_no."] = dt_i.MainProductionTaskNumber;
                        }

                        DatTool.RemoveSubstringFromDataTable(dt, "_OK");
                        mergeTable.Merge(dt);
                    }
                    string sheetname = dataInfoList[0].TestNumber + "_" + dataInfoList[0].PartNumber;
                    ExcelTool.ExportExcelWithDialog(mergeTable, sheetname, "MessData_" + dataInfoList[0].PartNumber);
                }
                else if (msDoub.IsChecked == true)
                {
                    for (int i = 0; i < dataTableList.Count; i++)
                    {
                        string psl = dataInfoList[i].TestSection.ToString();
                        string last = psl.Substring(psl.Length - 1);
                        if (ms_PSL.SelectedItem == null || ms_TestProg.SelectedItem == null)
                        {
                            MessageBox.Show("PSL and Test Program can not be null.\nPlease select an option.");
                            break;
                        }
                        else if (dataInfoList[i].TestNumber.Contains(ms_TestProg.SelectedItem.ToString()) && last == ms_PSL.Text)
                        {
                            DataTable dt = dataTableList[i];
                            Data_Infor dt_i = dataInfoList[i];

                            int moNoColumnIndex = dt.Columns.IndexOf("MO_no.");
                            if (moNoColumnIndex == -1)
                            {
                                DataColumn newColumn = new DataColumn("MO_no.", typeof(string));
                                dt.Columns.Add(newColumn);
                                dt.Columns["MO_no."].SetOrdinal(0);
                            }

                            foreach (DataRow row in dt.Rows)
                            {
                                row["MO_no."] = dt_i.MainProductionTaskNumber;
                            }

                            DatTool.RemoveSubstringFromDataTable(dt, "_OK");
                            mergeTable.Merge(dt);
                        }
                    }

                    string sheetname = ms_TestProg.Text + "_" + dataInfoList[0].PartNumber;
                    ExcelTool.ExportExcelWithDialog(mergeTable, sheetname, "MessData_" + dataInfoList[0].PartNumber);
                }
                else
                {
                    ExcelTool.ExportExcelWithListTable(dataTableList, "MessData");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public int ReadConfig(string path)
        {
            string[] config_content = File.ReadAllLines(path);
            foreach (string line in config_content)
            {
                string[] kytu;
                if (line.Contains("ODSSG_server"))
                {
                    kytu = line.Split('"');
                    Globalvariant.database = kytu[1].Trim();
                }
                if (line.Contains("ODSSG_user"))
                {
                    kytu = line.Split('"');
                    Globalvariant.user = kytu[1].Trim();
                }
                if (line.Contains("ODSSG_password"))
                {
                    kytu = line.Split('"');
                    Globalvariant.password = kytu[1].Trim();
                }
                if (line.Contains("ODSSG_service"))
                {
                    kytu = line.Split('"');
                    Globalvariant.service = kytu[1].Trim();
                }
                if (line.Contains("SMT="))
                {
                    kytu = line.Split('=');
                    Globalvariant.SMT = kytu[1].Trim();
                }
                if (line.Contains("Mainline_FA="))
                {
                    kytu = line.Split('=');
                    Globalvariant.FA = kytu[1].Trim();
                }
                if (line.Contains("Mainline_PA="))
                {
                    kytu = line.Split('=');
                    Globalvariant.PA = kytu[1].Trim();
                }
                if (line.Contains("HighVolt="))
                {
                    kytu = line.Split('=');
                    Globalvariant.Highvolt = kytu[1].Trim();
                }
            }
            if (Globalvariant.database == null || Globalvariant.SMT == null || Globalvariant.FA == null || Globalvariant.PA == null || Globalvariant.Highvolt == null)
            {
                return -1;
            }
            else
            {
                return 0;
            }
        }

        private void SelectTestProg()
        {
            try
            {
                #region Test Program
                HashSet<string> uniqueValues = new HashSet<string>();
                for (int i = 0; i < dataInfoList.Count; i++)
                {
                    uniqueValues.Add(dataInfoList[i].TestNumber);
                }
                ms_TestProg.ItemsSource = uniqueValues;
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SelectPSL()
        {
            try
            {
                #region Test Program
                HashSet<string> uniqueValues = new HashSet<string>();
                if (ms_TestProg.SelectedItem != null)
                {
                    for (int i = 0; i < dataInfoList.Count; i++)
                    {
                        if (dataInfoList[i].TestNumber == ms_TestProg.SelectedItem.ToString())
                        {
                            uniqueValues.Add(dataInfoList[i].TestSection.Substring(dataInfoList[i].TestSection.Length - 1));
                        }
                    }
                }
                ms_PSL.ItemsSource = uniqueValues;
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private async void CollectData(object sender, RoutedEventArgs e)
        {
            try
            {
                ClearData(sender, e);
                ProgressBox prog = new ProgressBox();
                prog.Show();

                UpdateProgressDelegate updateProgress = new UpdateProgressDelegate(prog.UpdateProgress);

                if (mes_Frdate.SelectedDate == null || mes_Todate.SelectedDate == null)
                {
                    MessageBox.Show("Please choose a correct time span for data.");
                    return;
                }
                if (msItem.Text == "")
                {
                    MessageBox.Show("Please choose an item.");
                    return;
                }
                int error = ReadConfig(PathReader.filePath);
                string messFolder = GetFolder();
                if (error != 0)
                {
                    MessageBox.Show("Missing config in Path_{Server}.ini file");
                    return;
                }
                updateProgress.Invoke(5, "5");
                await Task.Delay(10);
                DataTable table = new DataTable();
                try
                {
                    #region Connect Database and Create Input
                    using (OdbcConnection cnt = new OdbcConnection(Globalvariant.database))
                    {
                        cnt.Open();
                        if (cnt.State == ConnectionState.Open)
                        {
                            using (OdbcCommand cmd = new OdbcCommand())
                            {
                                cmd.Connection = cnt;
                                cmd.CommandType = CommandType.Text;

                                DateTime fr = mes_Frdate.SelectedDate ?? DateTime.MinValue;
                                DateTime to = mes_Todate.SelectedDate ?? DateTime.MaxValue;
                                string input = msItem.Text.Trim();
                                string fromdate = fr.ToString("dd-MMM-yyyy");
                                string todate = to.ToString("dd-MMM-yyyy");
                                updateProgress.Invoke(10, "10");
                                await Task.Delay(10);

                                string[] lines = File.ReadAllLines(PathReader.Movex);
                                foreach (string line in lines)
                                {
                                    if (line.Contains("CommandText_MO"))
                                    {
                                        string[] part = line.Split(new char[] { '=' }, 2);
                                        cmd.CommandText = part[1].ToString();
                                    }
                                }

                                cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();
                                cmd.Parameters.Add("@fromdate", OdbcType.Char).Value = fromdate.Trim();
                                cmd.Parameters.Add("@todate", OdbcType.Char).Value = todate.Trim();

                                using (OdbcDataAdapter data = new OdbcDataAdapter())
                                {
                                    data.SelectCommand = cmd;
                                    data.Fill(table);
                                }
                                updateProgress.Invoke(15, "15");
                                await Task.Delay(10);
                            }
                            if (table.Rows.Count == 0)
                            {
                                MessageBox.Show("Can not find MO list from Server\nPlease check item number or time span again");
                            }
                            cnt.Close();
                        }
                        else
                        {
                            MessageBox.Show("Can not connect to the database.");
                        }
                    }
                    #endregion

                    #region Copy file to local folder
                    DatTool.DeleteAllFiles(PathReader.boxplot_foldermes);

                    double x = 75.0 / table.Rows.Count;
                    double y = 15.0;
                    double z = 15.0;

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        string[] file = Directory.GetFiles(messFolder, "*" + table.Rows[i]["VHMFNO"].ToString() + "*");
                        if (file.Length != 0)
                        {
                            if (x + y - z >= 1.0)
                            {
                                z = x + y;
                                int intValue = (int)Math.Floor(z);

                                updateProgress.Invoke(intValue, intValue.ToString());
                                await Task.Delay(10);
                            }
                            y += x;

                            for (int j = 0; j < file.Length; j++)
                            {
                                DataTable addtable = new DataTable();
                                string[] filename = file[j].Split('\\');
                                string sourceFile = Path.Combine(messFolder, filename[filename.Length - 1]);
                                string copyFile = Path.Combine(PathReader.boxplot_foldermes, filename[filename.Length - 1]);

                                try
                                {
                                    if (File.Exists(sourceFile))
                                    {
                                        File.Copy(sourceFile, copyFile, true);
                                        DatTool.ReadFile(copyFile, addtable);
                                        mes_List.Items.Add(filename[filename.Length - 1]);
                                    }
                                    else
                                    {
                                        Console.WriteLine($"File not found: {sourceFile}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error copying file: {ex.Message}");
                                }
                            }
                        }
                    }
                    updateProgress.Invoke(100, "100");
                    await Task.Delay(10);
                    prog.Close();
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                SelectTestProg();
                SelectTestStep();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SelectTestStep()
        {
            try
            {
                ms_Step.Items.Clear();
                if (ms_TestProg.SelectedItem != null)
                {
                    string key = ms_TestProg.SelectedItem.ToString();
                    for (int i = 0; i < dataInfoList.Count; i++)
                    {
                        if (dataInfoList[i].TestNumber.Contains(key))
                        {
                            foreach (DataColumn column in dataTableList[i].Columns)
                            {
                                ms_Step.Items.Add(column.ColumnName);
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void MsStep_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (filePathList.Count != 0)
            {
                if (StepTable != null)
                {
                    StepTable.Columns.Clear();
                    StepTable.Rows.Clear();
                }

                string link = "";
                string key = ms_TestProg.SelectedItem.ToString();
                for (int i = 0; i < dataInfoList.Count; i++)
                {
                    if (dataInfoList[i].TestNumber.Contains(key))
                    {
                        link = filePathList[i];
                        break;
                    }
                }

                StreamReader reader = new StreamReader(link);
                StepTable = new DataTable();
                StepTable.Columns.Add("Step");
                StepTable.Columns.Add("Min");
                StepTable.Columns.Add("Max");
                StepTable.Columns.Add("Unit");

                ms_minv.Text = "-";
                ms_maxv.Text = "-";
                ms_unitv.Text = "-";

                string line;
                bool isDataSection = false;

                while ((line = reader.ReadLine()) != null)
                {
                    if (!isDataSection && line.Contains("Data="))
                    {
                        isDataSection = true;
                        break;
                    }

                    if (!isDataSection && line.Contains("Instruction="))
                    {
                        string[] values = line.Split('\t');
                        if (values.Length >= 7)
                        {
                            DataRow row = StepTable.NewRow();
                            row["Step"] = values[3].Trim();
                            row["Min"] = values[4].Trim();
                            row["Max"] = values[5].Trim();
                            if (values[6].Trim().Contains("�A"))
                            {
                                row["Unit"] = "µA";
                            }
                            else if (values[6].Trim().Contains("�") && !values[6].Trim().Contains("�A"))
                            {
                                row["Unit"] = "°C";
                            }
                            else
                            {
                                row["Unit"] = values[6].Trim();
                            }
                            StepTable.Rows.Add(row);
                        }
                    }
                }

                if (isDataSection && ms_Step.SelectedItem != null)
                {
                    foreach (DataRow row in StepTable.Rows)
                    {
                        if (ms_Step.SelectedItem.ToString() == row["Step"].ToString())
                        {
                            column_Name = ms_Step.SelectedItem.ToString();
                            Min_Value = row["Min"].ToString();
                            ms_minv.Text = Min_Value;
                            Max_Value = row["Max"].ToString();
                            ms_maxv.Text = Max_Value;
                            ms_unitv.Text = row["Unit"].ToString();
                        }
                    }
                }
            }
        }

        private void ProgSelect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectPSL();
            SelectTestStep();
        }

        private void Search_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && msPath.IsFocused)
            {
                OpenFiles(sender, e);
            }
            else if (e.Key == Key.Enter && msItem.IsFocused)
            {
                CollectData(sender, e);
            }
        }

        private void BoxPlotData()
        {
            #region Create table
            try
            {
                boxplot_dt.Clear();
                string item = dataInfoList[0].PartNumber;
                string selectPSL = "";
                string selectTestProg = "";
                string selectStep = "";

                if (ms_PSL.SelectedItem != null && ms_TestProg.SelectedItem != null && ms_Step.SelectedItem != null)
                {
                    selectPSL = ms_PSL.SelectedItem.ToString();
                    selectTestProg = ms_TestProg.SelectedItem.ToString();
                    selectStep = ms_Step.SelectedItem.ToString();
                }

                foreach (DataTable dataTable in dataTableList)
                {
                    boxplot_dt.Add(dataTable);
                }

                for (int i = 0; i < boxplot_dt.Count; i++)
                {
                    string psl = dataInfoList[i].TestSection.ToString();
                    string last = psl.Substring(psl.Length - 1);

                    if (dataInfoList[i].TestNumber.Contains(selectTestProg) && last == selectPSL)
                    {
                        if (ms_BoxAll.IsChecked == true)
                        {
                            for (int j = boxplot_dt[i].Rows.Count - 1; j >= 0; j--)
                            {
                                if (boxplot_dt[i].Rows[j]["Result"].ToString() == "D")
                                {
                                    boxplot_dt[i].Rows.RemoveAt(j);
                                }
                            }
                        }

                        DataTable filteredDataTable = new DataTable();
                        filteredDataTable.Columns.Add(selectStep);
                        foreach (DataRow row in boxplot_dt[i].Rows)
                        {
                            filteredDataTable.ImportRow(row);
                        }
                        BoxPlot.filter_tb.Add(filteredDataTable);
                        BoxPlot.mo_list.Add(dataInfoList[i].MainProductionTaskNumber);
                        BoxPlot.date.Add(dataInfoList[i].SystemDate);
                    }
                }
                BoxPlot.min_v = double.Parse(ms_minv.Text);
                BoxPlot.max_v = double.Parse(ms_maxv.Text);
                BoxPlot.median = (BoxPlot.min_v + BoxPlot.max_v) / 2;
                BoxPlot.BoxPlotTitle = $"BoxPlot of item: {item} -- Test Program: {ms_TestProg.SelectedItem}\nTest Step: {ms_Step.SelectedItem} [ {BoxPlot.min_v.ToString("N2")} {ms_unitv.Text} to {BoxPlot.max_v.ToString("N2")} {ms_unitv.Text} ]";
                LineChart.LineChartTitle = $"Line Chart of item: {item} -- Test Program: {ms_TestProg.SelectedItem}\nTest Step: {ms_Step.SelectedItem} [ {BoxPlot.min_v.ToString("N2")} {ms_unitv.Text} to {BoxPlot.max_v.ToString("N2")} {ms_unitv.Text} ]";
                BoxPlot.unit_v = ms_unitv.Text;

                for (int i = 0; i < BoxPlot.filter_tb.Count; i++)
                {
                    for (int x = 0; x < BoxPlot.filter_tb[i].Rows.Count; x++)
                    {
                        for (int y = 0; y < BoxPlot.filter_tb[i].Columns.Count; y++)
                        {
                            if (BoxPlot.filter_tb[i].Rows[x][y] != DBNull.Value && BoxPlot.filter_tb[i].Rows[x][y] is string)
                            {
                                string cellValue = (string)BoxPlot.filter_tb[i].Rows[x][y];
                                if (cellValue.Contains("_OK"))
                                {
                                    cellValue = cellValue.Replace("_OK", "");
                                    BoxPlot.filter_tb[i].Rows[x][y] = cellValue;
                                }
                            }
                        }
                    }

                    List<double> columnValues = BoxPlot.filter_tb[i].AsEnumerable().Select(row => Convert.ToDouble(row[selectStep])).ToList();
                    BoxPlotItem new_value = CalculateBoxPlotValues(columnValues, i);
                    BoxPlot.list_box.Add(new_value);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            #endregion
        }

        private BoxPlotItem CalculateBoxPlotValues(List<double> data, int i)
        {
            data.Sort();
            double lowerWhisker = data.Min();
            double boxBottom = CalculateMedian(data.Take(data.Count / 2).ToList());
            double median = CalculateMedian(data);
            double boxTop = CalculateMedian(data.Skip((data.Count + 1) / 2).ToList());
            double upperWhisker = data.Max();

            return new BoxPlotItem(i, lowerWhisker, boxBottom, median, boxTop, upperWhisker);
        }

        private double CalculateMedian(List<double> values)
        {
            int count = values.Count;
            if (count % 2 == 0)
            {
                return (values[count / 2 - 1] + values[count / 2]) / 2.0;
            }
            else
            {
                return values[count / 2];
            }
        }

        private void CreateBP(object sender, RoutedEventArgs e)
        {
            BoxPlot.list_box.Clear();
            BoxPlot.filter_tb.Clear();
            BoxPlotData();

            BoxPlot boxplot = new BoxPlot();
            boxplot.Show();
        }

        private void CreateLC(object sender, RoutedEventArgs e)
        {
            BoxPlot.list_box.Clear();
            BoxPlot.filter_tb.Clear();
            BoxPlotData();

            LineChart lineChart = new LineChart();
            lineChart.Show();
        }

        private void ListViewSelect(object sender, SelectionChangedEventArgs e)
        {
            string folderPath = PathReader.boxplot_foldermes;
            string[] files = Directory.GetFiles(folderPath);
            string select = mes_List.SelectedItem.ToString();
            string[] filename = select.Split('_');

            if (mes_List.SelectedItem != null)
            {
                foreach (string filePath in files)
                {
                    try
                    {
                        string fileContent = File.ReadAllText(filePath);
                        if (fileContent.Contains(filename[1]))
                        {
                            Process.Start(filePath);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void PSL_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (ms_PSL.SelectedItem != null && ms_TestProg.SelectedItem != null)
                {
                    string selected = ms_PSL.SelectedItem.ToString();
                    for (int i = 0; i < dataInfoList.Count; i++)
                    {
                        string comp = dataInfoList[i].TestSection.Substring(dataInfoList[i].TestSection.Length - 1);
                        if (ms_PSL.SelectedItem.ToString() == comp && dataInfoList[i].TestNumber.Contains(ms_TestProg.SelectedItem.ToString()))
                        {
                            ms_Section.Text = dataInfoList[i].TestSectionName;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void OpenFolder(object sender, RoutedEventArgs e)
        {
            Process.Start(PathReader.boxplot_foldermes);
        }
    }
}