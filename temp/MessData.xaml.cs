using MnS.lib;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using System.Data.Odbc;
using System.Windows.Controls;

namespace MnS
{
    /// <summary>
    /// 
    /// </summary>
    public class Data_Infor
    {
        #region MO Data
        public string ProductionTaskNumber { get; set; }
        public string ReferencedProductionTaskNumber { get; set; }
        public string MainProductionTaskNumber { get; set; }
        public string PartNumber { get; set; }
        public string SerialNumber { get; set; }
        public string TestNumber { get; set; }
        public string OPNumber { get; set; }
        public string PersonalNumber { get; set; }
        public string TestSection { get; set; }
        public string TestSectionName { get; set; }
        public string SoftwareVersion { get; set; }
        public string SystemDate { get; set; }
        public string TestProgramDate { get; set; }
        public string[] CalibrationNumber { get; set; }
        public string[] FixtureNumbers { get; set; }
        public string ArrayLayout { get; set; }
        public string Remark1 { get; set; }
        public string Remark2 { get; set; }
        public string Prodstatus { get; set; }
        public string TestStartMode { get; set; }
        public string[] RCIEDMSettings { get; set; }
        public string RCIPARSetting { get; set; }
        public string RCIPROGSetting { get; set; }
        public string RCIPPSSetting { get; set; }
        public string RCILCSetting { get; set; }
        public string BASE { get; set; }
        public string[] Firmware_1 { get; set; }
        public string[] Firmware_2 { get; set; }
        public string[] Firmware_3 { get; set; }
        public string[] Firmware_4 { get; set; }
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="lines"></param>
        public Data_Infor(string[] lines)
        {
            ProcessLines(lines);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="lines"></param>
        private void ProcessLines(string[] lines)
        {
            foreach (var line in lines)
            {
                if (line.Contains("="))
                {
                    string[] tokens = line.Split('=');
                    #region Create MO Data
                    switch (tokens[0])
                    {
                        case "Production task number":
                            ProductionTaskNumber = tokens[1].Replace("\t", "");
                            break;
                        case "Referenced production task number":
                            ReferencedProductionTaskNumber = tokens[1].Replace("\t", "");
                            break;
                        case "Main production task number":
                            MainProductionTaskNumber = tokens[1].Replace("\t", "");
                            break;
                        case "Part number":
                            PartNumber = tokens[1].Replace("\t", "");
                            break;
                        case "Serial number":
                            SerialNumber = tokens[1].Replace("\t", "");
                            break;
                        case "Test number":
                            TestNumber = tokens[1].Replace("\t", "");
                            break;
                        case "OP number":
                            OPNumber = tokens[1].Replace("\t", "");
                            break;
                        case "Personal number":
                            PersonalNumber = tokens[1].Replace("\t", "");
                            break;
                        case "Test section":
                            TestSection = tokens[1].Replace("\t", "");
                            break;
                        case "Test section name":
                            TestSectionName = tokens[1].Replace("\t", "");
                            break;
                        case "Software version":
                            SoftwareVersion = tokens[1].Replace("\t", "");
                            break;
                        case "System date":
                            SystemDate = tokens[1].Replace("\t", "");
                            break;
                        case "Test program date":
                            TestProgramDate = tokens[1].Replace("\t", "");
                            break;
                        case "Calibration number":
                            tokens[1] = tokens[1].Replace("\t", "");
                            CalibrationNumber = tokens[1].Split(';');
                            break;
                        case "Fixture number":
                            tokens[1] = tokens[1].Replace("\t", "");
                            tokens[1] = tokens[1].Replace(";", "");
                            FixtureNumbers = tokens[1].Split(' ');
                            break;
                        case "Array layout":
                            ArrayLayout = tokens[1].Replace("\t", "");
                            break;
                        case "Remark1":
                            Remark1 = tokens[1].Replace("\t", "");
                            break;
                        case "Remark2":
                            Remark2 = tokens[1].Replace("\t", "");
                            break;
                        case "Prodstatus":
                            Prodstatus = tokens[1].Replace("\t", "");
                            break;
                        case "Test Start mode":
                            TestStartMode = tokens[1].Replace("\t", "");
                            break;
                        case "RCI EDM Setting":
                            tokens[1] = tokens[1].Replace("\t", " ");
                            RCIEDMSettings = tokens[1].Split(';');
                            break;
                        case "RCI PAR Setting":
                            RCIPARSetting = tokens[1].Replace("\t", "");
                            break;
                        case "RCI PROG Setting":
                            RCIPROGSetting = tokens[1].Replace("\t", "");
                            break;
                        case "RCI PPS Setting":
                            RCIPPSSetting = tokens[1].Replace("\t", "");
                            break;
                        case "RCI LC Setting":
                            RCILCSetting = tokens[1].Replace("\t", "");
                            break;
                        case "firmware_1":
                            tokens[1] = tokens[1].Trim();
                            Firmware_1 = tokens[1].Split('\t');
                            break;
                        case "firmware_2":
                            tokens[1] = tokens[1].Trim();
                            Firmware_2 = tokens[1].Split('\t');
                            break;
                        case "firmware_3":
                            tokens[1] = tokens[1].Trim();
                            Firmware_3 = tokens[1].Split('\t');
                            break;
                        case "firmware_4":
                            tokens[1] = tokens[1].Trim();
                            Firmware_4 = tokens[1].Split('\t');
                            break;
                        case "BASE":
                            BASE = tokens[1].Replace("\t", "");
                            break;
                    }
                    #endregion
                }
            }
        }
    }

    /// <summary>
    /// Interaction logic for MessData.xaml
    /// </summary>
    public partial class MessData : Window
    {
        #region
        DatabaseConnection.Movex movex = new DatabaseConnection.Movex();
        public static DataTable temp_Data;
        public DataTable step_Table;
        public List<DataTable> temp_Table = new List<DataTable>();
        public List<Data_Infor> dataInfoList = new List<Data_Infor>();
        Filehandle.File file = new Filehandle.File();
        OdbcConnection cnt;

        public string[] selectedFiles;
        public string[] files;
        public string filePath;
        public string file_Path;
        public string column_Name;
        public string Min_Value;
        public string Max_Value;
        public static string Unit;
        public static string messdatafolder = null;
        #endregion

        /// <summary>
        /// 
        /// </summary>
        public static class Globalvariant
        {
            public static string database, user, password, service, SMT, PA, FA, Highvolt = null;
        }

        /// <summary>
        /// 
        /// </summary>
        public MessData()
        {
            InitializeComponent();
            ms_Step.SelectionChanged += MsStep_SelectionChanged;
        }

        /// <summary>
        /// 
        /// </summary>
        public void OpenFiles(object sender, RoutedEventArgs e)
        {
            temp_Table.Clear();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (msMulti.IsChecked == true)
            {
                openFileDialog.Multiselect = true;
                dataInfoList.Clear();
            }
            openFileDialog.Filter = "DAT Files|*.DAT|All Files|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                if (msMulti.IsChecked == false)
                {
                    file_Path = openFileDialog.FileName;
                    string fileName = Path.GetFileName(file_Path);
                    msPath.Text = fileName;
                }

                foreach (string file in openFileDialog.FileNames)
                {
                    DataTable dataTable = new DataTable(GetDataTableNameFromFileName(file));
                    filePath = file;
                    OpenAndCopyFiles(filePath);
                    string fileName = Path.GetFileName(filePath);
                    mes_List.Items.Clear();
                    mes_List.Items.Add(fileName);
                    ReadFile(dataTable, filePath);
                    temp_Table.Add(temp_Data);
                }

                bool differencesFound = false;

                if (msMulti.IsChecked == false && temp_Table.Count == 1)
                {
                    ComboBoxTool.DisplayComboBox(temp_Data, ms_Step);
                }
                else if (msMulti.IsChecked == true && temp_Table.Count > 1)
                {
                    foreach (Data_Infor data_Infor in dataInfoList)
                    {
                        if (!data_Infor.PartNumber.Equals(dataInfoList[0].PartNumber) || !data_Infor.TestNumber.Equals(dataInfoList[0].TestNumber))
                        {
                            differencesFound = true;
                            break;
                        }
                    }
                }

                if (differencesFound)
                {
                    ms_minv.Text = "-";
                    ms_maxv.Text = "-";
                    ms_unitv.Text = "-";
                    mes_SgGrid2.IsEnabled = false;
                    MessageBox.Show("Differences Item data!. Can not create chart.\nPlease export Excel file!.");
                }
                else
                {
                    foreach (DataTable dataTable in temp_Table)
                    {
                        temp_Data.Merge(dataTable);
                    }

                    ComboBoxTool.DisplayComboBox(temp_Data, ms_Step);
                    mes_SgGrid2.IsEnabled = true;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public void ReadFile(DataTable dataTable, string filepath)
        {
            StreamReader reader = new StreamReader(filepath, Encoding.UTF8);
            string[] lines = File.ReadAllLines(filepath);
            Data_Infor dataInfo = new Data_Infor(lines);
            dataInfoList.Add(dataInfo);

            dataTable.Columns.Add("Tester");
            dataTable.Columns.Add("No.");
            dataTable.Columns.Add("A3");
            dataTable.Columns.Add("Result");
            dataTable.Columns.Add("A5");

            string line;
            while ((line = reader.ReadLine()) != null)
            {
                if (line.Contains("Instruction="))
                {
                    string[] values = line.Split('\t');
                    for (int i = 0; i < values.Length; i++)
                    {
                        values[i] = values[i].Trim();
                    }

                    if (!dataTable.Columns.Contains(values[3].Trim()))
                    {
                        dataTable.Columns.Add(values[3].Trim());
                    }
                }

                try
                {
                    if (line.Contains("Data="))
                    {
                        string[] values = line.Split('\t');
                        for (int i = 0; i < values.Length; i++)
                        {
                            values[i] = values[i].Trim();
                        }

                        if (values.Length != dataTable.Columns.Count)
                        {
                            DataRow row = dataTable.NewRow();
                            for (int i = 0; i < dataTable.Columns.Count; i++)
                            {
                                row[i] = values[i + 1];
                            }
                            dataTable.Rows.Add(row);
                        }
                        else
                        {
                            MessageBox.Show("Data is not correct, please check again!");
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Data is not correct, please check again!");
                }
            }

            dataTable.DefaultView.Sort = "Tester ASC";
            SwapDataTableColumns(dataTable, "A3", "SerNr");
            dataTable.Columns.Remove("A3");
            dataTable.Columns.Remove("A5");

            temp_Data = dataTable;
            _ = dataTable.DefaultView.ToTable();

            reader.Close();
        }

        /// <summary>
        /// 
        /// </summary>
        private void OpenAndCopyFiles(string filepath)
        {
            string tempFolder = PathReader.boxplot_folder;
            if (!Directory.Exists(tempFolder))
            {
                Directory.CreateDirectory(tempFolder);
            }

            string fileName = Path.GetFileName(filepath);
            string destinationPath = Path.Combine(tempFolder, fileName);
            File.Copy(filepath, destinationPath, true);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private string GetDataTableNameFromFileName(string fileName)
        {
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            int lengthLimit = 11;
            string name = nameWithoutExtension.Substring(nameWithoutExtension.Length - lengthLimit);
            string name_1 = name.Substring(0, name.Length - 1);;
            return name_1;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="cl_1"></param>
        /// <param name="cl_2"></param>
        public void SwapDataTableColumns(DataTable dt, string cl_1, string cl_2)
        {
            if (dt.Columns.Contains(cl_1) && dt.Columns.Contains(cl_2))
            {
                int column1Ordinal = dt.Columns[cl_1].Ordinal;
                int column2Ordinal = dt.Columns[cl_2].Ordinal;
                dt.Columns[cl_1].SetOrdinal(column2Ordinal);
                dt.Columns[cl_2].SetOrdinal(column1Ordinal);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public int ReadConfig(string path)
        {
            string[] config_content = File.ReadAllLines(path);
            _ = new string[2];
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
                return -1;
            else
                return 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Collect_Data(object sender, RoutedEventArgs e)
        {
            Clear_Data(sender, e);
            int error = ReadConfig(PathReader.filePath);
            if (error != 0)
            {
                MessageBox.Show("Missing config in Path_IDN.ini file");
                return;
            }
            cnt = new OdbcConnection();
            DataTable table = new DataTable();

            try
            {
                using (OdbcConnection cnt = new OdbcConnection(Globalvariant.database))
                {
                    cnt.Open();

                    if (cnt.State == ConnectionState.Open)
                    {
                        using (OdbcCommand rt_cmd = new OdbcCommand())
                        {
                            rt_cmd.Connection = cnt;
                            rt_cmd.CommandType = CommandType.Text;

                            DateTime fr = mes_Frdate.SelectedDate ?? DateTime.MinValue;
                            DateTime to = mes_Todate.SelectedDate ?? DateTime.MaxValue;
                            string input = msItem.Text.Trim();
                            string fromdate = fr.ToString("dd-MMM-yyyy");
                            string todate = to.ToString("dd-MMM-yyyy");

                            string[] lines = File.ReadAllLines(PathReader.Movex);
                            foreach (string line in lines)
                            {
                                if (line.Contains("CommandText_MO"))
                                {
                                    string[] part = line.Split(new char[] { '=' }, 2);
                                    rt_cmd.CommandText = part[1].ToString();
                                }
                            }

                            rt_cmd.Parameters.Add("@input", OdbcType.Char).Value = input.Trim();
                            rt_cmd.Parameters.Add("@fromdate", OdbcType.Char).Value = fromdate.Trim();
                            rt_cmd.Parameters.Add("@todate", OdbcType.Char).Value = todate.Trim();

                            using (OdbcDataAdapter rt_data = new OdbcDataAdapter())
                            {
                                rt_data.SelectCommand = rt_cmd;
                                rt_data.Fill(table);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Can not connect to the database.");
                    }
                }

                if (ms_SMT.IsChecked == true)
                {
                    messdatafolder = Globalvariant.SMT;
                }
                if (ms_FA.IsChecked == true)
                {
                    messdatafolder = Globalvariant.FA;
                }
                if (ms_PA.IsChecked == true)
                {
                    messdatafolder = Globalvariant.PA;
                }
                if (ms_HV.IsChecked == true)
                {
                    messdatafolder = Globalvariant.Highvolt;
                }

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    string[] file = Directory.GetFiles(messdatafolder, "*" + table.Rows[i]["VHMFNO"].ToString() + "*");
                    if (file.Length != 0)
                    {
                        for (int j = 0; j < file.Length; j++)
                        {
                            string[] filename = file[j].Split('\\');
                            string sourceFile = Path.Combine(messdatafolder, filename[filename.Length - 1]);
                            string copyFile = Path.Combine(PathReader.boxplot_foldermes, filename[filename.Length - 1]);

                            try
                            {
                                if (File.Exists(sourceFile))
                                {
                                    File.Copy(sourceFile, copyFile, true);
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

                List<string> List_PSL = new List<string>();
                List<string> List_TestProg = new List<string>();
                files = Directory.GetFiles(PathReader.boxplot_foldermes, "*.DAT");
                foreach (string file_ in files)
                {
                    bool PSL_datontai = false;
                    bool TestProg_datontai = false;
                    string PSL = file_.Substring(file_.Length - 5, 1);
                    string[] steps = file_.Split('\\');
                    string[] step_1 = steps[steps.Length - 1].Split('_');
                    string TestProg = step_1[0];
                    for (int i = 0; i < List_PSL.Count; i++)
                        if (PSL.Trim() == List_PSL[i])
                            PSL_datontai = true;
                    for (int j = 0; j < List_TestProg.Count; j++)
                        if (TestProg.Trim() == List_TestProg[j])
                            TestProg_datontai = true;
                    if (PSL_datontai == false)
                    {
                        List_PSL.Add(PSL.Trim());
                    }

                    if (TestProg_datontai == false)
                        List_TestProg.Add(TestProg.Trim());
                }
                for (int i = 0; i < List_PSL.Count; i++)
                {
                    ms_PSL.Items.Add(List_PSL[i]);
                }
                for (int i = 0; i < List_TestProg.Count; i++)
                {
                    ms_TestProg.Items.Add(List_TestProg[i]);
                }

                ms_BoxAll.IsEnabled = false;
                cnt.Close();
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Connect failed"))
                    MessageBox.Show("Connection to database fail\nPlease check your internet connection or your database config");
                else if (ex.Message.Contains("Data source name not found"))
                    MessageBox.Show("Connection to database fail\nPlease check your database config in Path_{Server}.ini file");
                else
                    MessageBox.Show("Error:" + ex);
            };
            if (table.Rows.Count == 0)
                MessageBox.Show("Can not find MO list from Server\nPlease check item number or time span again");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Clear_Data(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(PathReader.boxplot_folder))
            {
                Directory.CreateDirectory(PathReader.boxplot_folder);
            }
            else
            {
                string[] folder = Directory.GetFiles(PathReader.boxplot_folder, "*.DAT");
                foreach (string file in folder)
                {
                    File.Delete(file);
                }
            }

            if (!Directory.Exists(PathReader.boxplot_foldermes))
            {
                Directory.CreateDirectory(PathReader.boxplot_foldermes);
            }
            else
            {
                string[] folder_mes = Directory.GetFiles(PathReader.boxplot_foldermes, "*.DAT");
                foreach (string file_mes in folder_mes)
                {
                    File.Delete(file_mes);
                }
            }

            if(!Directory.Exists(PathReader.boxplot_foldercp))
            {
                Directory.CreateDirectory(PathReader.boxplot_foldercp);
            }
            else
            {
                string[] folder_cp = Directory.GetFiles(PathReader.boxplot_foldercp, "*.DAT");
                foreach (string file_cp in folder_cp)
                {
                    File.Delete(file_cp);
                }
            }
            mes_List.Items.Clear();
            ms_BoxAll.IsEnabled = false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Convert_Data(object sender, RoutedEventArgs e)
        {
            try
            {
                if (msPath.Text != "" && msPath.Text != "View list files")
                {
                    DataTable dt = new DataTable();
                    ReadFile(dt, file_Path);

                    ExcelTool.ExportExcelWithName(temp_Data, "Mess_Data");
                }
                else if (mes_List.Items != null && temp_Table.Count != 0)
                {
                    ExcelTool.ExportExcelWithListTable(temp_Table, "Mess_Data.xlsx");
                }
                Process.Start(ExcelTool.excel_filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Multi_Checked(object sender, RoutedEventArgs e)
        {
            ms_Step.Items.Clear();
            msPath.Text = "View list files";
            mes_SgGrid1.IsEnabled = true;
            mes_SgGrid2.IsEnabled = false;
            mes_DouGrid1.IsEnabled = false;
            mes_DouGrid2.IsEnabled = false;
            mes_DouGrid3.IsEnabled = true;
            ms_Box.IsEnabled = true;
            Clear_Data(sender, e);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Multi_Unchecked(object sender, RoutedEventArgs e)
        {
            msPath.Text = "";
            msDoub.IsChecked = false;
            mes_SgGrid1.IsEnabled = true;
            mes_SgGrid2.IsEnabled = true;
            mes_DouGrid1.IsEnabled = false;
            mes_DouGrid2.IsEnabled = false;
            mes_DouGrid3.IsEnabled = false;
            ms_Box.IsEnabled = false;
            Clear_Data(sender, e);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Doub_Checked(object sender, RoutedEventArgs e)
        {
            msMulti.IsChecked = true;
            mes_SgGrid1.IsEnabled = false;
            mes_SgGrid2.IsEnabled = false;
            mes_DouGrid1.IsEnabled = true;
            mes_DouGrid2.IsEnabled = true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Doub_Unchecked(object sender, RoutedEventArgs e)
        {
            mes_SgGrid1.IsEnabled = true;
            mes_SgGrid2.IsEnabled = false;
            mes_DouGrid1.IsEnabled = false;
            mes_DouGrid2.IsEnabled = false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void MsStep_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            StreamReader reader = new StreamReader(filePath, Encoding.UTF8);

            DataTable step_Data = new DataTable();
            step_Data.Columns.Add("Step");
            step_Data.Columns.Add("Min");
            step_Data.Columns.Add("Max");
            step_Data.Columns.Add("Unit");

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
                        DataRow row = step_Data.NewRow();
                        row["Step"] = values[3].Trim();
                        row["Min"] = values[4].Trim();
                        row["Max"] = values[5].Trim();
                        if (values[6].Trim().Contains("�A"))
                        {
                            row["Unit"] = "µA";
                        }
                        else
                        {
                            row["Unit"] = values[6].Trim();
                        }
                        step_Data.Rows.Add(row);
                    }
                }
            }
            step_Table = step_Data;

            if (isDataSection && ms_Step.SelectedItem != null)
            {
                foreach (DataRow row in step_Table.Rows)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Create_Graph(object sender, RoutedEventArgs e)
        {
            if (ms_Step.SelectedItem != null)
            {
                DataTable value_Table = temp_Data.Copy();

                double lower = double.Parse(Min_Value);
                double upper = double.Parse(Max_Value);
                double average = (lower + upper) / 2;
                string chose_column = column_Name;

                bool minExists = value_Table.Columns.Contains("Min");
                bool maxExists = value_Table.Columns.Contains("Max");
                bool averageExists = value_Table.Columns.Contains("Average");

                if (!minExists)
                {
                    value_Table.Columns.Add("Min", typeof(double));
                }
                if (!maxExists)
                {
                    value_Table.Columns.Add("Max", typeof(double));
                }
                if (!averageExists)
                {
                    value_Table.Columns.Add("Average", typeof(double));
                }

                foreach (DataRow row in value_Table.Rows)
                {
                    foreach (DataColumn column in value_Table.Columns)
                    {
                        if (row[column] is string @string && @string.Contains("_OK"))
                        {
                            row[column] = @string.Replace("_OK", "");
                        }

                        if (double.TryParse(row[column].ToString(), out double cellValue))
                        {
                            row[column] = cellValue;
                        }
                    }

                    row["Min"] = lower;
                    row["Max"] = upper;
                    row["Average"] = average;
                }

                List<DataRow> rowsToRemove = new List<DataRow>();
                foreach (DataRow row in value_Table.Rows)
                {
                    if (row["Result"].ToString().Contains("D") && ms_Pass.IsChecked == true)
                    {
                        rowsToRemove.Add(row);
                    }
                }

                foreach (DataRow rowToRemove in rowsToRemove)
                {
                    value_Table.Rows.Remove(rowToRemove);
                }

                List<string> columnsToRemove = new List<string>();
                foreach (DataColumn column in value_Table.Columns)
                {
                    if (!column.ColumnName.Equals(chose_column) && !column.ColumnName.Equals("Min") && !column.ColumnName.Equals("Max") && !column.ColumnName.Equals("Average"))
                    {
                        columnsToRemove.Add(column.ColumnName);
                    }
                }

                foreach (string columnToRemove in columnsToRemove)
                {
                    value_Table.Columns.Remove(columnToRemove);
                }

                ExcelTool.ExportExcelWithPath(PathReader.temp_chart, value_Table);

                value_Table.Rows.Clear();
                value_Table.Columns.Clear();
                Process.Start(PathReader.Line_chart);
            }
            else
            {
                MessageBox.Show("Please select a test step to create chart.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Create_Box(object sender, RoutedEventArgs e)
        {
            if (ms_Step.SelectedItem != null)
            {
                string file_config = PathReader.boxplot_folder + "test_step.ini";
                file.createFile(file_config);
                using (StreamWriter wstream = new StreamWriter(file_config))
                {
                    wstream.WriteLine(ms_Step.SelectedItem.ToString());
                    if (ms_Pass.IsChecked == true)
                    {
                        wstream.WriteLine("GOOD");
                    }
                    else
                    {
                        wstream.WriteLine("ALL");
                    }
                }
                Process.Start(PathReader.Box_Plot);
            }
            else
            {
                MessageBox.Show("Please select a test step to create chart.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MsProg_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                string[] folder = Directory.GetFiles(PathReader.boxplot_folder, "*.DAT");
                foreach (string file in folder)
                {
                    File.Delete(file);
                }

                string[] files = Directory.GetFiles(PathReader.boxplot_foldermes);
                foreach (string filePath in files)
                {
                    if (Path.GetFileName(filePath).Contains(ms_TestProg.SelectedItem.ToString()) && Path.GetFileName(filePath).Contains(ms_PSL.SelectedItem.ToString() + ".DAT"))
                    {
                        string destinationFilePath = Path.Combine(PathReader.boxplot_folder, Path.GetFileName(filePath));
                        File.Copy(filePath, destinationFilePath, true);
                    }
                }
                ms_BoxAll.IsEnabled = true;
            }
            catch (Exception ex)
            {
                if (ms_TestProg.SelectedItem != null && ms_PSL.SelectedItem != null)
                {
                    MessageBox.Show("PSL and Test Program can not be null, please select.");
                }
                else
                {
                    MessageBox.Show("Error: " + ex);
                }
            }

            DataTable temp = new DataTable();
            string[] folder_temp = Directory.GetFiles(PathReader.boxplot_folder, "*.DAT");
            ReadFile(temp, folder_temp[0]);
            ComboBoxTool.DisplayComboBox(temp_Data, ms_TestStep);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Search_CreateBox(object sender, RoutedEventArgs e)
        {
            if (ms_TestStep.SelectedItem != null)
            {
                string file_config = PathReader.boxplot_folder + "\\test_step.ini";
                file.createFile(file_config);
                using (StreamWriter wstream = new StreamWriter(file_config))
                {
                    wstream.WriteLine(ms_TestStep.SelectedItem.ToString());
                    if (ms_Pass_Copy.IsChecked == true)
                    {
                        wstream.WriteLine("GOOD");
                    }
                    else
                    {
                        wstream.WriteLine("ALL");
                    }
                }
                Process.Start(PathReader.Box_Plot);
                ms_CVAll.IsEnabled = true;
            }
            else
            {
                MessageBox.Show("Please select a test step to create chart.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Convert_DataAll(object sender, RoutedEventArgs e)
        {
            try
            {
                string[] folder = Directory.GetFiles(PathReader.boxplot_foldercp, "*.DAT");
                List<DataTable> temp_Tables = new List<DataTable>();

                foreach (string file in folder)
                {
                    DataTable dataTable = new DataTable(GetDataTableNameFromFileName(file));
                    OpenAndCopyFiles(file);
                    string fileName = Path.GetFileName(file);
                    ReadFile(dataTable, file);
                    temp_Tables.Add(dataTable);
                }

                ExcelTool.ExportExcelWithListTable(temp_Tables, "temp_data.xlsx");

                MessageBox.Show("Convert all files to excel (.xlsx) successfully.");
                Process.Start(ExcelTool.excel_filePath);

                ms_CVAll.IsEnabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        #region CheckBox
        /// <summary>
        /// 
        /// </summary>
        /// <param name="check"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="chosebox"></param>
        /// <param name="hidenboxes"></param>
        public static void CheckAndUncheck(CheckBox chosebox, List<CheckBox> hidenboxes)
        {
            if (chosebox == null || hidenboxes == null)
                return;
            hidenboxes.Remove(chosebox);
            foreach (CheckBox check in hidenboxes)
            {
                if (check != null && check.IsChecked == true)
                    check.IsChecked = false;
            }
            hidenboxes.Add(chosebox);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void HV_Checked(object sender, RoutedEventArgs e)
        {
            Mescheck(ms_HV);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FA_Checked(object sender, RoutedEventArgs e)
        {
            Mescheck(ms_FA);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PA_Checked(object sender, RoutedEventArgs e)
        {
            Mescheck(ms_PA);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SMT_Checked(object sender, RoutedEventArgs e)
        {
            Mescheck(ms_SMT);
        }
        #endregion
    }
}