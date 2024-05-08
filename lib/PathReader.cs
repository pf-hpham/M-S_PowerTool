using System;
using System.IO;
using System.Windows;

namespace MnS.lib
{
    public static class PathReader
    {
        static PathReader()
        {
            Read_Server();
            Read_Parameter();
        }

        public static void Read_Server()
        {
            Config_Check();
            string filepath = "C:\\M+S_Server\\Path_Server.ini";
            string movex = "C:\\M+S_Server\\Movex_Server.ini";
            string filesource;
            string movexsource;

            try
            {
                string[] SEA_region = File.ReadAllLines(server);
                foreach (string sea_line in SEA_region)
                {
                    if (sea_line.Contains("Server="))
                    {
                        string[] part = sea_line.Split(new char[] { '=' }, 2);
                        if (part[1].ToString() == "VIE")
                        {
                            filesource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Path_VIE.ini";
                            Path_Check(filepath, filesource);
                            movexsource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Movex_VIE.ini";
                            Path_Check(movex, movexsource);
                            filePath = filepath;
                            Movex = movex;
                        }
                        else if (part[1].ToString() == "SGP")
                        {
                            filesource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Path_SGP.ini";
                            Path_Check(filepath, filesource);
                            movexsource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Movex_SGP.ini";
                            Path_Check(movex, movexsource);
                            filePath = filepath;
                            Movex = movex;
                        }
                        else if (part[1].ToString() == "IDN")
                        {
                            filesource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Path_IDN.ini";
                            Path_Check(filepath, filesource);
                            movexsource = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Movex_IDN.ini";
                            Path_Check(movex, movexsource);
                            filePath = filepath;
                            Movex = movex;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\nThe system will automatically create the connection for the first time.c\nClick OK to continue.");
            }
        }

        private static void Config_Check()
        {
            string config_path = "C:\\M+S_Server\\Config.ini";
            string source = "\\\\pfvn-netapp1\\files\\87-Maintenance-Services-SEA\\_Public\\100-M+S_PowerTool\\config\\Config.ini";
            Path_Check(config_path, source);
            config = config_path;
        }

        public static void Path_Check(string destinate, string link_source)
        {
            try
            {
                if (!File.Exists(destinate))
                {
                    File.Create(destinate).Close();
                    File.Copy(link_source, destinate, true);
                    MessageBox.Show($"Created new {destinate} for the first use succesfully.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\nThe system will automatically create the connection for the first time.\nClick OK to continue.");
            }
        }

        private static void Read_Parameter()
        {
            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);
                #region VietNam
                foreach (string line in lines)
                {
                    if (line.Contains("="))
                    {
                        string[] parts = line.Split(new char[] { '=' }, 2);

                        if (parts.Length == 2)
                        {
                            string first = parts[0].Trim();
                            string second = parts[1].Trim().Trim('"');
                            if (first == "CAL_link")
                            {
                                CAL_link = second;
                            }
                            else if (first == "PM_link")
                            {
                                PM_link = second;
                            }
                            else if (first == "Quote_folder")
                            {
                                Quote_folder = second;
                            }
                            else if (first == "Image_user")
                            {
                                Image_user = second;
                            }
                            else if (first == "Image_device")
                            {
                                Image_device = second;
                            }
                            else if (first == "Log_file")
                            {
                                Log_file = second;
                            }
                            else if (first == "Log_data")
                            {
                                Log_data = second;
                            }
                            else if (first == "CALmsp_file")
                            {
                                CALmsp_file = second;
                            }
                            else if (first == "CALrecord_file")
                            {
                                CALrecord_file = second;
                            }
                            else if (first == "EDM_link")
                            {
                                EDM_link = second;
                            }
                            else if (first == "ODSSG_server")
                            {
                                ODSSG_server = second;
                            }
                            else if (first == "EDM_server")
                            {
                                EDM_server = second;
                            }
                            else if (first == "Guide_link")
                            {
                                Guide_link = second;
                            }
                            else if (first == "EDM_link")
                            {
                                EDM_link = second;
                            }
                            else if (first == "Teamspace")
                            {
                                Teamspace = second;
                            }
                            else if (first == "Techmaster")
                            {
                                Techmaster = second;
                            }
                            else if (first == "EzPortal")
                            {
                                EzPortal = second;
                            }
                            else if (first == "Misumi")
                            {
                                Misumi = second;
                            }
                            else if (first == "RSComponent")
                            {
                                RSComponent = second;
                            }
                            else if (first == "GGSearch")
                            {
                                GGSearch = second;
                            }
                            else if (first == "Line_chart")
                            {
                                Line_chart = second;
                            }
                            else if (first == "temp_chart")
                            {
                                temp_chart = second;
                            }
                            else if (first == "SMT")
                            {
                                SMT = second;
                            }
                            else if (first == "Main_Line")
                            {
                                Main_Line = second;
                            }
                            else if (first == "High_Volt")
                            {
                                High_Volt = second;
                            }
                            else if (first == "Box_Plot")
                            {
                                Box_Plot = second;
                            }
                            else if (first == "boxplot_folder")
                            {
                                boxplot_folder = second;
                            }
                            else if (first == "boxplot_foldermes")
                            {
                                boxplot_foldermes = second;
                            }
                            else if (first == "boxplot_foldercp")
                            {
                                boxplot_foldercp = second;
                            }
                            else if (first == "boxplot_config")
                            {
                                boxplot_config = second;
                            }
                        }
                    }
                }
                #endregion
            }
        }

        #region public variable
        public static string Download_link = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads\\";
        public static string MNSP_link = "Data Source=1904VND028\\SQLEXPRESS;Initial Catalog=MNSP;User ID=common;Password=Pfvn123";
        public static string DVM_link = "Data Source=1904VND028\\SQLEXPRESS;Initial Catalog=DVM;User ID=common;Password=Pfvn123";
        public static string server = "C:\\M+S_Server\\Server.ini";
        public static string config { get; set; }
        public static string filePath { get; set; }
        public static string Movex { get; set; }
        public static string CAL_link { get; set; }
        public static string PM_link { get; set; }
        public static string Image_user { get; set; }
        public static string Image_device { get; set; }
        public static string Quote_folder { get; set; }
        public static string Log_file { get; set; }
        public static string Log_data { get; set; }
        public static string CALmsp_file { get; set; }
        public static string CALrecord_file { get; set; }
        public static string EDM_link { get; set; }
        public static string ODSSG_server { get; set; }
        public static string EDM_server { get; set; }
        public static string Guide_link { get; set; }
        public static string Teamspace { get; set; }
        public static string Techmaster { get; set; }
        public static string EzPortal { get; set; }
        public static string Misumi { get; set; }
        public static string RSComponent { get; set; }
        public static string GGSearch { get; set; }
        public static string Line_chart { get; set; }
        public static string Box_Plot { get; set; }
        public static string temp_chart { get; set; }
        public static string SMT { get; set; }
        public static string Main_Line { get; set; }
        public static string High_Volt { get; set; }
        public static string boxplot_folder { get; set; }
        public static string boxplot_foldermes { get; set; }
        public static string boxplot_foldercp { get; set; }
        public static string boxplot_config { get; set; }
        #endregion
    }
}