using MnS.lib;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace MnS
{
    public partial class EDM : Window
    {
        #region variable
        public delegate void UpdateProgressDelegate(int value, string prog);

        public static string Input;
        public static string Infor;
        public static string ReDate;
        private OdbcConnection cnt;

        public static DataTable dt;
        public static DataTable ref_dt;
        public static DataTable cmt_dt;

        public static string CommandText_13;
        public static string CommandText_14;
        public static string CommandText_15;
        #endregion

        public EDM()
        {
            UserLogTool.UserData("Using EDM function");
            InitializeComponent();
            ReadCommandLine();
            edm_input.Focus();
            edm_input.SelectAll();
        }

        private void ReadCommandLine()
        {
            string[] lines = File.ReadAllLines(PathReader.Movex);
            foreach (string line in lines)
            {
                if (line.Contains("CommandText"))
                {
                    string[] part = line.Split(new char[] { '=' }, 2);
                    if (part[0].ToString() == "CommandText_13")
                    {
                        CommandText_13 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_14")
                    {
                        CommandText_14 = part[1].ToString();
                    }
                    else if (part[0].ToString() == "CommandText_15")
                    {
                        CommandText_15 = part[1].ToString();
                    }
                }
            }
        }

        private OdbcCommand ConnectDataBase(string CommandText)
        {
            cnt = new OdbcConnection
            {
                ConnectionString = PathReader.EDM_server
            };
            cnt.Open();

            OdbcCommand cmd = new OdbcCommand
            {
                Connection = cnt,
                CommandType = CommandType.Text,
                CommandText = CommandText
            };

            return cmd;
        }

        public void GetData(string input)
        {
            if (dt != null)
            {
                dt.Clear();
                edm_index.Items.Clear();
            }

            dt = new DataTable();
            string command = CommandText_13;
            OdbcCommand cmd = ConnectDataBase(command);
            cmd.Parameters.Add("@input", OdbcType.Char).Value = "%" + input.Trim() + "%";
            OdbcDataAdapter getDataData = new OdbcDataAdapter();
            getDataData.SelectCommand = cmd;
            getDataData.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    edm_index.Items.Add(row["DOCNUMBER"].ToString());
                }
                edm_index.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("Document is not exits.");
            }
        }

        public async void Infor_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
                ProgressBox prog = new ProgressBox();
                prog.Show();

                UpdateProgressDelegate updateProgress = new UpdateProgressDelegate(prog.UpdateProgress);
                updateProgress.Invoke(10, "10");
                await Task.Delay(10);

                foreach (DataRow row in dt.Rows)
                {
                    if (row["DOCNUMBER"].ToString() == edm_index.SelectedItem.ToString())
                    {
                        Infor = row["INFO"].ToString() + "\n" + row["INFO1"].ToString();
                        ReDate = row["RELEASEDATE"].ToString();
                        break;
                    }
                }
                updateProgress.Invoke(30, "30");
                await Task.Delay(10);

                edm_readt.Text = ReDate;
                edm_des.Text = Infor;

                Refer_Function();
                updateProgress.Invoke(60, "60");
                await Task.Delay(10);

                His_Function();
                updateProgress.Invoke(90, "90");
                await Task.Delay(10);
                updateProgress.Invoke(100, "100");
                prog.Close();
            }
        }

        public void Refer_Function()
        {
            if (ref_dt != null)
            {
                ref_dt.Clear();
            }

            ref_dt = new DataTable();
            string command = CommandText_14;
            OdbcCommand cmd = ConnectDataBase(command);
            cmd.Parameters.Add("@input", OdbcType.Char).Value = edm_index.SelectedItem.ToString();
            OdbcDataAdapter getDataData = new OdbcDataAdapter();
            getDataData.SelectCommand = cmd;
            getDataData.Fill(ref_dt);

            DateTimeFormat.DatetimeFormat(ref_dt, edm_ref, "A");
        }

        public void His_Function()
        {
            if (cmt_dt != null)
            {
                cmt_dt.Clear();
            }

            cmt_dt = new DataTable();
            string command = CommandText_15;
            OdbcCommand cmd = ConnectDataBase(command);
            cmd.Parameters.Add("@input", OdbcType.Char).Value = edm_index.SelectedItem.ToString();
            OdbcDataAdapter getDataData = new OdbcDataAdapter();
            getDataData.SelectCommand = cmd;
            getDataData.Fill(cmt_dt);

            DateTimeFormat.DatetimeFormat(cmt_dt, edm_his, "A");
        }

        private void EdmInput_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Input = edm_input.Text.ToUpper();
                GetData(Input);
            }
        }

        private void EDM_Click(object sender, MouseButtonEventArgs e)
        {
            DataRowView dataRow = (DataRowView)edm_ref.SelectedItem;
            int index = edm_ref.CurrentCell.Column.DisplayIndex;
            string cellValue = dataRow.Row.ItemArray[index].ToString();

            if (cellValue == " " || cellValue == "" || cellValue == null)
            {
                MessageBox.Show("Cell value is null.");
            }
            else
            {
                if (index == 1 && edm_view.IsChecked == false)
                {
                    Input = cellValue;
                    edm_input.Text = cellValue;
                    GetData(Input);
                }
                else if (edm_view.IsChecked == true)
                {
                    Process.Start(PathReader.EDM_link + cellValue);
                }
            }
        }
    }
}