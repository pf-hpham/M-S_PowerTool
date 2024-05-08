using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows;

namespace MnS.lib
{
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

        public Data_Infor(string[] lines)
        {
            foreach (var line in lines)
            {
                if (line.Contains("="))
                {
                    string[] tokens = line.Split('=');
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
                }
            }
        }
    }

    public static class DatTool
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="dataTable"></param>
        public static void ReadFile(string filepath, DataTable dataTable)
        {
            StreamReader reader = new StreamReader(filepath, Encoding.UTF8);
            string[] lines = File.ReadAllLines(filepath);
            Data_Infor data_Infor = new Data_Infor(lines);
            string key = data_Infor.TestNumber;
            string dataName = key + "_" + GetNameFromFileName(filepath);
            dataTable.TableName = dataName;

            dataTable.Columns.Add("Tester");
            dataTable.Columns.Add("No.");
            dataTable.Columns.Add("A3");
            dataTable.Columns.Add("Result");
            dataTable.Columns.Add("A5");

            string line;
            while ((line = reader.ReadLine()) != null)
            {
                try
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
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex);
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
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex);
                }
            }

            dataTable.DefaultView.Sort = "Tester ASC";
            SwapDataTableColumns(dataTable, "A3", "SerNr");
            dataTable.Columns.Remove("A3");
            dataTable.Columns.Remove("A5");
            dataTable.DefaultView.ToTable();

            MessData.dataInfoList.Add(data_Infor);
            MessData.dataTableList.Add(dataTable);
            MessData.filePathList.Add(filepath);
            reader.Close();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="cl_1"></param>
        /// <param name="cl_2"></param>
        public static void SwapDataTableColumns(DataTable dt, string cl_1, string cl_2)
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
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string GetNameFromFileName(string fileName)
        {
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            int lengthLimit = 11;
            string name = nameWithoutExtension.Substring(nameWithoutExtension.Length - lengthLimit);
            string name_1 = name.Substring(0, name.Length - 1); ;
            return name_1;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="destinationFolder"></param>
        public static void CopyFilesToFolder(List<string> files, string destinationFolder)
        {
            if (files != null && files.Count > 0)
            {
                foreach (string sourceFile in files)
                {
                    string fileName = Path.GetFileName(sourceFile);
                    string destinationPath = Path.Combine(destinationFolder, fileName);
                    File.Copy(sourceFile, destinationPath, true);
                }
            }
            else
            {
                MessageBox.Show("No files selected.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="folderPath"></param>
        public static void DeleteAllFiles(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                string[] datFiles = Directory.GetFiles(folderPath, "*.DAT");
                foreach (string datFile in datFiles)
                {
                    File.Delete(datFile);
                }
            }
            else
            {
                Console.WriteLine("File folder is not exits.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="substring"></param>
        public static void RemoveSubstringFromDataTable(DataTable dataTable, string substring)
        {
            foreach (DataRow row in dataTable.Rows)
            {
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    if (row[i] != DBNull.Value && row[i] is string @string)
                    {
                        row[i] = @string.Replace(substring, string.Empty);
                    }
                }
            }
        }
    }
}