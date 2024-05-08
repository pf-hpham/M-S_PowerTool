using System;
using System.Data;
using System.Linq;
using LiveCharts;
using MnS.lib;
using LiveCharts.Wpf;
using System.Windows.Media;
using System.Data.SqlClient;
using System.Windows.Controls;
using System.Collections.Generic;

namespace MnS
{
    public partial class PM_Dashboard : UserControl
    {
        public PM_Dashboard()
        {
            InitializeComponent();
            Location_Chart();
            Cart_Chart();
            ImExp_Chart();
            Store_Chart();
        }

        private void Location_Chart()
        {
            DataTable LocData = SQLDataTool.QueryUserData("SELECT Change_Level FROM Gage_Master", new List<SqlParameter>(), PathReader.PM_link);

            DataTable changeLevelCounts = new DataTable();
            changeLevelCounts.Columns.Add("Change_Level", typeof(string));
            changeLevelCounts.Columns.Add("Count", typeof(int));

            var changeLevelGroups = LocData.AsEnumerable()
                .GroupBy(row => row.Field<string>("Change_Level"))
                .Select(group => new
                {
                    ChangeLevel = group.Key,
                    Count = group.Count()
                })
                .OrderBy(group => group.ChangeLevel ?? "No Defined");

            foreach (var group in changeLevelGroups)
            {
                DataRow newRow = changeLevelCounts.NewRow();
                newRow["Change_Level"] = group.ChangeLevel ?? "No Defined";
                newRow["Count"] = group.Count;
                changeLevelCounts.Rows.Add(newRow);
            }

            SeriesCollection locationSeriesCollection = new SeriesCollection();
            foreach (DataRow row in changeLevelCounts.Rows)
            {
                string changeLevel = row["Change_Level"].ToString();
                int count = Convert.ToInt32(row["Count"]);

                locationSeriesCollection.Add(new PieSeries
                {
                    Title = changeLevel,
                    Values = new ChartValues<double> { count },
                    DataLabels = true,
                    LabelPoint = point => $"{count}",
                });
            }

            Location_qttchart.Series = locationSeriesCollection;
        }

        private void Store_Chart()
        {
            DataTable LocData = SQLDataTool.QueryUserData("SELECT GM_Owner, User_Defined FROM Gage_Master", new List<SqlParameter>(), PathReader.PM_link);

            int msTmCount = LocData.AsEnumerable().Count(row => row.Field<string>("GM_Owner") == "M+S TM");
            int msPeCount = LocData.AsEnumerable().Count(row => row.Field<string>("GM_Owner") == "M+S P&E");
            int msFtCount = LocData.AsEnumerable().Count(row => row.Field<string>("GM_Owner") == "M+S F&T");

            SeriesCollection storeSeriesCollection = new SeriesCollection
            {
                new PieSeries
                {
                    Title = "MnS TM",
                    Values = new ChartValues<int> { msTmCount },
                    DataLabels = true,
                    LabelPoint = point => $"{msTmCount}",
                },
                new PieSeries
                {
                    Title = "MnS PE",
                    Values = new ChartValues<int> { msPeCount },
                    DataLabels = true,
                    LabelPoint = point => $"{msPeCount}",
                },
                new PieSeries
                {
                    Title = "MnS FT",
                    Values = new ChartValues<int> { msFtCount },
                    DataLabels = true,
                    LabelPoint = point => $"{msFtCount}",
                }
            };

            int total_Quantity = LocData.Rows.Count;
            str_quan.Text = total_Quantity.ToString();
            Store_qttchart.Series = storeSeriesCollection;
        }

        private void Cart_Chart()
        {
            DataTable PMTable = SQLDataTool.QueryUserData("SELECT Current_Location, Status FROM Gage_Master", new List<SqlParameter>(), PathReader.PM_link);

            List<string> uniqueLocations = PMTable.AsEnumerable()
                .Select(row => row.Field<string>("Current_Location"))
                .Distinct()
                .OrderBy(location => location)
                .ToList();

            DataTable columnData = new DataTable();
            columnData.Columns.Add("Location", typeof(string));
            columnData.Columns.Add("Count", typeof(int));

            DataTable lineData = new DataTable();
            lineData.Columns.Add("Location", typeof(string));
            lineData.Columns.Add("Status", typeof(int));

            foreach (string location in uniqueLocations)
            {
                int status;
                int loc_count = PMTable.AsEnumerable().Count(row => row.Field<string>("Current_Location") == $"{location}");
                int stt_count = PMTable.AsEnumerable().Count(row => row.Field<string>("Current_Location") == $"{location}" && (int.TryParse(row["Status"].ToString(), out status) && (status == 1 || status == 4)));

                DataRow columnDataRow = columnData.NewRow();
                columnDataRow["Location"] = location;
                columnDataRow["Count"] = loc_count;
                columnData.Rows.Add(columnDataRow);

                DataRow lineDataRow = lineData.NewRow();
                lineDataRow["Location"] = location;
                lineDataRow["Status"] = stt_count;
                lineData.Rows.Add(lineDataRow);
            }

            ColumnSeries columnSeries = new ColumnSeries
            {
                Title = "Total Devices",
                Values = new ChartValues<int>(columnData.AsEnumerable().Select(row => row.Field<int>("Count"))),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString(),
                FontSize = 8,
                ScalesYAt = 0
            };

            LineSeries lineSeries = new LineSeries
            {
                Title = "Need Maintenance",
                Values = new ChartValues<int>(lineData.AsEnumerable().Select(row => row.Field<int>("Status"))),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString(),
                FontSize = 8,
                ScalesYAt = 0
            };

            quantity_chart.AxisY.Clear();
            quantity_chart.AxisY.Add(new Axis
            {
                Title = "Total",
                LabelFormatter = value => value.ToString(),
                Position = AxisPosition.LeftBottom,
                Foreground = Brushes.Black,
                FontSize = 10,
                MinValue = 0
            });

            string[] locationLabels = uniqueLocations.ToArray();
            quantity_chart.AxisX.Clear();
            quantity_chart.AxisX.Add(new Axis
            {
                Foreground = Brushes.Black,
                FontSize = 10,
                Labels = locationLabels,
                LabelsRotation = -20,
            });

            quantity_chart.Series.Clear();
            quantity_chart.Series.Add(columnSeries);
            quantity_chart.Series.Add(lineSeries);
        }

        private void ImExp_Chart()
        {
            DataTable PMData = SQLDataTool.QueryUserData("SELECT Next_Due_Date, Last_Calibration_Date FROM Gage_Master", new List<SqlParameter>(), PathReader.PM_link);
            int currentYear = DateTime.Now.Year;

            CartesianChart imExp_chart = new CartesianChart();

            imexp_chart.AxisX.Clear();
            imexp_chart.AxisX.Add(new Axis
            {
                Title = "Chart of PM schedule for months in " + currentYear,
                Labels = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" },
                Foreground = Brushes.Black
            });

            imexp_chart.AxisY.Clear();
            imexp_chart.AxisY.Add(new Axis
            {
                LabelFormatter = value => value.ToString(),
                Foreground = Brushes.Black,
                FontSize = 10
            });

            SeriesCollection seriesCollection = new SeriesCollection();
            LineSeries importSeries = new LineSeries
            {
                Title = "The total number of devices not yet maintained in the month",
                Values = new ChartValues<int>(),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString(),
                FontSize = 10
            };

            LineSeries exportSeries = new LineSeries
            {
                Title = "The total number of devices maintained during the month",
                Values = new ChartValues<int>(),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString(),
                ScalesYAt = 0,
                FontSize = 10
            };

            EnumerableRowCollection<DataRow> filteredImportData = PMData.AsEnumerable().Where(row => !row.IsNull("Next_Due_Date") && row.Field<DateTime>("Next_Due_Date").Year == currentYear);
            EnumerableRowCollection<DataRow> filteredExportData = PMData.AsEnumerable().Where(row => !row.IsNull("Last_Calibration_Date") && row.Field<DateTime>("Last_Calibration_Date").Year == currentYear);

            var importDataByMonth = filteredImportData.GroupBy(row => row.Field<DateTime>("Next_Due_Date").Month).Select(group => new
            {
                Month = group.Key,
                TotalImportQuantity = group.Count()
            });

            var exportDataByMonth = filteredExportData.GroupBy(row => row.Field<DateTime>("Last_Calibration_Date").Month).Select(group => new
            {
                Month = group.Key,
                TotalExportQuantity = group.Count()
            });

            for (int i = 1; i <= 12; i++)
            {
                int importQuantity = importDataByMonth.FirstOrDefault(item => item.Month == i)?.TotalImportQuantity ?? 0;
                int exportQuantity = exportDataByMonth.FirstOrDefault(item => item.Month == i)?.TotalExportQuantity ?? 0;

                importSeries.Values.Add(importQuantity);
                exportSeries.Values.Add(exportQuantity);
            }

            imexp_chart.Series = seriesCollection;
            imexp_chart.Series.Add(importSeries);
            imexp_chart.Series.Add(exportSeries);
        }
    }
}