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
    public partial class CAL_Dashboard : UserControl
    {
        /// <summary>
        /// 
        /// </summary>
        public CAL_Dashboard()
        {
            InitializeComponent();
            Location_Chart();
            Cart_Chart();
            ImExp_Chart();
            Store_Chart();
        }

        /// <summary>
        /// 
        /// </summary>
        private void Location_Chart()
        {
            DataTable LocData = SQLDataTool.QueryUserData("SELECT Change_Level FROM Gage_Master", new List<SqlParameter>(), PathReader.CAL_link);

            DataTable changeLevelCounts = new DataTable();
            changeLevelCounts.Columns.Add("Change_Level", typeof(string));
            changeLevelCounts.Columns.Add("Count", typeof(int));

            var changeLevelGroups = LocData.AsEnumerable()
                .GroupBy(row => row.Field<string>("Change_Level"))
                .Select(group => new
                {
                    ChangeLevel = group.Key,
                    Count = group.Count()
                });

            foreach (var group in changeLevelGroups)
            {
                DataRow newRow = changeLevelCounts.NewRow();
                newRow["Change_Level"] = group.ChangeLevel;
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

        /// <summary>
        /// 
        /// </summary>
        private void Store_Chart()
        {
            DataTable LocData = SQLDataTool.QueryUserData("SELECT GM_Owner, Calibrated_By, User_Defined FROM Gage_Master", new List<SqlParameter>(), PathReader.CAL_link);
            int techmasterCount = LocData.AsEnumerable().Count(row => row.Field<string>("Calibrated_By") == "TECHMASTER");
            int tamsuiCount = LocData.AsEnumerable().Count(row => row.Field<string>("Calibrated_By") == "TAMSUI");
            int internalCount = LocData.AsEnumerable()
                .Count(row => row.Field<string>("Calibrated_By") != "TECHMASTER" &&
                              row.Field<string>("Calibrated_By") != "TAMSUI" &&
                              row.Field<string>("User_Defined") == "IN-HOUSE");

            int externalCount = LocData.Rows.Count - (techmasterCount + tamsuiCount + internalCount);

            SeriesCollection storeSeriesCollection = new SeriesCollection
            {
                new PieSeries
                {
                    Title = "TECHMASTER",
                    Values = new ChartValues<double> { techmasterCount },
                    DataLabels = true,
                    LabelPoint = point => $"{techmasterCount}",
                },
                new PieSeries
                {
                    Title = "TAMSUI",
                    Values = new ChartValues<double> { tamsuiCount },
                    DataLabels = true,
                    LabelPoint = point => $"{tamsuiCount}",
                },
                new PieSeries
                {
                    Title = "OTHER VENDOR",
                    Values = new ChartValues<double> { externalCount },
                    DataLabels = true,
                    LabelPoint = point => $"{externalCount}",
                },
                new PieSeries
                {
                    Title = "INTERNAL",
                    Values = new ChartValues<double> { internalCount },
                    DataLabels = true,
                    LabelPoint = point => $"{internalCount}",
                }
            };

            int total_Quantity = LocData.Rows.Count;
            str_quan.Text = total_Quantity.ToString();

            int msTmCount = LocData.AsEnumerable().Count(row => row.Field<string>("GM_Owner") == "M+S TM");
            int msPeCount = LocData.AsEnumerable().Count(row => row.Field<string>("GM_Owner") == "M+S P&E");

            tm_quan.Text = msTmCount.ToString();
            pe_quan.Text = msPeCount.ToString();

            Store_qttchart.Series = storeSeriesCollection;
        }

        /// <summary>
        /// 
        /// </summary>
        private void Cart_Chart()
        {
            DataTable CALTable = SQLDataTool.QueryUserData("SELECT Current_Location, Status FROM Gage_Master", new List<SqlParameter>(), PathReader.CAL_link);

            List<string> uniqueLocations = CALTable.AsEnumerable()
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
                int loc_count = CALTable.AsEnumerable().Count(row => row.Field<string>("Current_Location") == $"{location}");
                int stt_count = CALTable.AsEnumerable().Count(row => row.Field<string>("Current_Location") == $"{location}" && (int.TryParse(row["Status"].ToString(), out status) && (status == 1 || status == 4)));

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
                Title = "Total Device",
                Values = new ChartValues<int>(columnData.AsEnumerable().Select(row => row.Field<int>("Count"))),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString(),
                FontSize = 8,
                ScalesYAt = 0
            };

            LineSeries lineSeries = new LineSeries
            {
                Title = "Need Calibration",
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

        /// <summary>
        /// 
        /// </summary>
        private void ImExp_Chart()
        {
            DataTable CALData = SQLDataTool.QueryUserData("SELECT Next_Due_Date, Last_Calibration_Date FROM Gage_Master", new List<SqlParameter>(), PathReader.CAL_link);
            int currentYear = DateTime.Now.Year;

            CartesianChart imExp_chart = new CartesianChart();

            imexp_chart.AxisX.Clear();
            imexp_chart.AxisX.Add(new Axis
            {
                Title = "Chart of Calibration schedule for months in " + currentYear,
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
                Title = "The total number of devices not yet calibrated this month",
                Values = new ChartValues<int>(),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString(),
                FontSize = 10
            };

            LineSeries exportSeries = new LineSeries
            {
                Title = "The total number of devices calibrated this month (include lost, change status...)",
                Values = new ChartValues<int>(),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString(),
                ScalesYAt = 0,
                FontSize = 10
            };

            EnumerableRowCollection<DataRow> filteredImportData = CALData.AsEnumerable().Where(row => !row.IsNull("Next_Due_Date") && row.Field<DateTime>("Next_Due_Date").Year == currentYear);
            EnumerableRowCollection<DataRow> filteredExportData = CALData.AsEnumerable().Where(row => !row.IsNull("Last_Calibration_Date") && row.Field<DateTime>("Last_Calibration_Date").Year == currentYear);

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