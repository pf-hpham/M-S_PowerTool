using System;
using System.Data;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace MnS.lib
{
    public delegate void UpdateProgressDelegate(int value);

    public static class ProgressBarTool
    {
        public static async Task ProgressBarAsync(DataTable table, Action<double> progressUpdater)
        {
            double progress = 0;
            try
            {
                await Task.Run(() =>
                {
                    int rowCount = table.Rows.Count;
                    for (int i = 0; i < rowCount; i++)
                    {
                        progress = (i + 1) * 100.0 / rowCount;
                        progressUpdater(progress);
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during processing: " + ex.Message);
            }
        }

        public static void UpdateProgressBar(ProgressBar progressBar, double value)
        {
            progressBar.Dispatcher.Invoke(() =>
            {
                progressBar.Value = value;
            });
        }

        public static void CreateProgress()
        {
            
        }
    }
}
