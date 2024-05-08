using MnS.lib;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;

namespace MnS
{
    public partial class PM_Tab : UserControl
    {
        public static DataTable PMTable = new DataTable();

        public PM_Tab()
        {
            UserLogTool.UserData("Using Maintenance function");
            InitializeComponent();
        }

        private async void LoadData_Click(object sender, RoutedEventArgs e)
        {
            loadingProg.Visibility = Visibility.Visible;
            loadingProg.Value = 0;

            DataTable PM_Table = SQLDataTool.QueryUserData("SELECT * FROM Gage_Master", new List<SqlParameter>(), PathReader.PM_link);
            PMTable = PM_Table;

            await ProgressBarTool.ProgressBarAsync(PMTable, (progress) =>
            {
                ProgressBarTool.UpdateProgressBar(loadingProg, progress);
            });
            loadingProg.Visibility = Visibility.Hidden;

            DataContext = new ModelView.PMModelView(PMTable);
            MessageBox.Show("Loaded Database successfully!");
        }

        private void PMexport_Click(object sender, RoutedEventArgs e)
        {
            ExcelTool.ExportExcelWithDialog(PMTable, "PM", "PM_export_data");
        }

        private void PMSend_Click(object sender, RoutedEventArgs e)
        {
            if (PMTable != null)
            {
                OutlookTool.SendEmailWithExcelAttachment(PMTable);
            }
            else
            {
                MessageBox.Show("None of Data to Send.");
            }
        }
    }
}