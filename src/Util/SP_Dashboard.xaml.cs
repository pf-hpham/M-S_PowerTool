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
    public partial class SP_Dashboard : UserControl
    {
        /// <summary>
        /// 
        /// </summary>
        public SP_Dashboard()
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
        public string[] Labels { get; set; }

        /// <summary>
        /// 
        /// </summary>
        private void Location_Chart()
        {
            DataTable LocData = SQLDataTool.QueryUserData("SELECT TPU, Quantity FROM Device_List", new List<SqlParameter>(), PathReader.MNSP_link);
            DataTable uniqueTPUTotalQuantity = new DataTable();
            uniqueTPUTotalQuantity.Columns.Add("TPU", typeof(string));
            uniqueTPUTotalQuantity.Columns.Add("TotalQuantity", typeof(int));
            IEnumerable<(string TPU, int TotalQuantity)> query = from row in LocData.AsEnumerable() group row by row.Field<string>("TPU") into tpuGroup select (TPU: tpuGroup.Key, TotalQuantity: tpuGroup.Sum(row => row.Field<int>("Quantity")));

            foreach ((string TPU, int TotalQuantity) in query)
            {
                DataRow newRow = uniqueTPUTotalQuantity.NewRow();
                newRow["TPU"] = TPU;
                newRow["TotalQuantity"] = TotalQuantity;
                uniqueTPUTotalQuantity.Rows.Add(newRow);
            }

            SeriesCollection locationSeriesCollection = new SeriesCollection();
            foreach (DataRow row in uniqueTPUTotalQuantity.Rows)
            {
                string tpu = row["TPU"].ToString();
                int totalQuantity = Convert.ToInt32(row["TotalQuantity"]);

                locationSeriesCollection.Add(new PieSeries
                {
                    Title = tpu,
                    Values = new ChartValues<double> { totalQuantity },
                    DataLabels = true,
                    LabelPoint = point => $"{totalQuantity}",
                });
            }

            int total_Quantity = LocData.AsEnumerable().Sum(row => row.Field<int>("Quantity"));
            total_quan.Text = total_Quantity.ToString();

            Location_qttchart.Series = locationSeriesCollection;
        }

        /// <summary>
        /// 
        /// </summary>
        private void Store_Chart()
        {
            DataTable LocData = SQLDataTool.QueryUserData("SELECT Store_loc, Quantity FROM Device_List", new List<SqlParameter>(), PathReader.MNSP_link);
            DataTable uniqueStoreTotalQuantity = new DataTable();
            uniqueStoreTotalQuantity.Columns.Add("Store_loc", typeof(string));
            uniqueStoreTotalQuantity.Columns.Add("TotalQuantity", typeof(int));
            IEnumerable<(string Store_loc, int TotalQuantity)> query = from row in LocData.AsEnumerable() group row by row.Field<string>("Store_loc") into storeGroup select (Store_loc: storeGroup.Key, TotalQuantity: storeGroup.Sum(row => row.Field<int>("Quantity")));

            foreach ((string Store_loc, int TotalQuantity) in query)
            {
                DataRow newRow = uniqueStoreTotalQuantity.NewRow();
                newRow["Store_loc"] = Store_loc;
                newRow["TotalQuantity"] = TotalQuantity;
                uniqueStoreTotalQuantity.Rows.Add(newRow);
            }

            SeriesCollection storeSeriesCollection = new SeriesCollection();
            foreach (DataRow row in uniqueStoreTotalQuantity.Rows)
            {
                string storeLoc = row["Store_loc"].ToString();
                int totalQuantity = Convert.ToInt32(row["TotalQuantity"]);

                storeSeriesCollection.Add(new PieSeries
                {
                    Title = storeLoc,
                    Values = new ChartValues<double> { totalQuantity },
                    DataLabels = true,
                    LabelPoint = point => $"{totalQuantity}",
                });
            }

            int total_Quantity = LocData.AsEnumerable().Sum(row => row.Field<int>("Quantity"));
            str_quan.Text = total_Quantity.ToString();

            Store_qttchart.Series = storeSeriesCollection;
        }

        /// <summary>
        /// 
        /// </summary>
        private void Cart_Chart()
        {
            DataTable CartData = SQLDataTool.QueryUserData("SELECT Location, Quantity, Price, Currency FROM Device_List", new List<SqlParameter>(), PathReader.MNSP_link);
            DataTable PriData = SQLDataTool.QueryUserData("SELECT Currency_code, Conversion_rate FROM Currency", new List<SqlParameter>(), PathReader.MNSP_link);

            DataTable TT_Quantity_Table = new DataTable();
            TT_Quantity_Table.Columns.Add("Location", typeof(string));
            TT_Quantity_Table.Columns.Add("TT_Quantity", typeof(int));

            IEnumerable<(string Location, int TT_Quantity)> queryTT = from row in CartData.AsEnumerable() group row by row.Field<string>("Location") into locationGroup select (Location: locationGroup.Key, TT_Quantity: locationGroup.Sum(row => row.Field<int>("Quantity")));

            foreach ((string Location, int TT_Quantity) in queryTT)
            {
                DataRow newRow = TT_Quantity_Table.NewRow();
                newRow["Location"] = Location;
                newRow["TT_Quantity"] = TT_Quantity;
                TT_Quantity_Table.Rows.Add(newRow);
            }

            DataTable Total_Prices_Table = new DataTable();
            Total_Prices_Table.Columns.Add("Location", typeof(string));
            Total_Prices_Table.Columns.Add("Total_Prices", typeof(double));

            foreach (DataRow row in CartData.Rows)
            {
                string location = row.Field<string>("Location");
                int quantity = row.Field<int>("Quantity");
                double price = row.Field<double>("Price");
                string currency = row.Field<string>("Currency");

                double conversionRate = PriData.AsEnumerable().Where(r => r.Field<string>("Currency_code") == currency).Select(r => r.Field<double>("Conversion_rate")).FirstOrDefault();

                double totalPrice = quantity * price / conversionRate;

                DataRow existingRow = Total_Prices_Table.AsEnumerable().FirstOrDefault(r => r.Field<string>("Location") == location);

                if (existingRow != null)
                {
                    existingRow["Total_Prices"] = existingRow.Field<double>("Total_Prices") + totalPrice;
                }
                else
                {
                    DataRow newRow = Total_Prices_Table.NewRow();
                    newRow["Location"] = location;
                    newRow["Total_Prices"] = totalPrice;
                    Total_Prices_Table.Rows.Add(newRow);
                }
            }

            DataView cartDataView = TT_Quantity_Table.DefaultView.ToTable().DefaultView;
            ColumnSeries columnSeries = new ColumnSeries
            {
                Title = "Total quantity = ",
                Values = new ChartValues<int>(),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString() + " pcs",
                FontSize = 8,
                ScalesYAt = 0
            };

            foreach (DataRowView rowView in cartDataView)
            {
                int ttQuantity = Convert.ToInt32(rowView["TT_Quantity"]);
                columnSeries.Values.Add(ttQuantity);
            }

            SeriesCollection seriesCollection = new SeriesCollection
            {
                columnSeries
            };

            LineSeries lineSeries = new LineSeries
            {
                Title = "Total prices = ",
                Values = new ChartValues<double>(),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString("F2") + " USD",
                FontSize = 8,
                ScalesYAt = 1
            };
            quantity_chart.AxisY.Clear();
            quantity_chart.AxisY.Add(new Axis
            {
                Title = "Quantity and Total SP value of Lines",
                LabelFormatter = value => value.ToString(),
                Position = AxisPosition.LeftBottom,
                Foreground = Brushes.Black,
                FontSize = 10
            });
            quantity_chart.AxisY.Add(new Axis
            {
                LabelFormatter = value => value.ToString(),
                Position = AxisPosition.RightTop,
                Foreground = Brushes.Black,
                FontSize = 10
            });

            foreach (DataRow row in Total_Prices_Table.Rows)
            {
                double totalPrices = Convert.ToDouble(row["Total_Prices"]);
                lineSeries.Values.Add(totalPrices);
            }

            SeriesCollection lineCollection = new SeriesCollection
            {
                lineSeries
            };

            Labels = Total_Prices_Table.AsEnumerable().Select(row => row.Field<string>("Location")).ToArray();

            quantity_chart.AxisX.Clear();
            quantity_chart.AxisX.Add(new Axis
            {
                Foreground = Brushes.Black,
                FontSize = 10,
                Labels = Labels,
                LabelsRotation = -20,
            });

            quantity_chart.Series = seriesCollection;
            quantity_chart.Series.Add(lineSeries);
        }

        /// <summary>
        /// 
        /// </summary>
        private void ImExp_Chart()
        {
            DataTable DevData = SQLDataTool.QueryUserData("SELECT Register_date, Quantity FROM Device_List", new List<SqlParameter>(), PathReader.MNSP_link);
            DataTable ImpData = SQLDataTool.QueryUserData("SELECT Import_date, Import_quantity FROM Import_History", new List<SqlParameter>(), PathReader.MNSP_link);
            DataTable ExpData = SQLDataTool.QueryUserData("SELECT Export_date, Export_quantity FROM Export_History", new List<SqlParameter>(), PathReader.MNSP_link);
            int currentYear = DateTime.Now.Year;

            CartesianChart imExp_chart = new CartesianChart();

            imexp_chart.AxisX.Clear();
            imexp_chart.AxisX.Add(new Axis
            {
                Title = "Chart of Import-Export frequency for months in " + currentYear,
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
                Title = "Import Monthly",
                Values = new ChartValues<int>(),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString(),
                FontSize = 10
            };

            LineSeries exportSeries = new LineSeries
            {
                Title = "Export Monthly",
                Values = new ChartValues<int>(),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString(),
                ScalesYAt = 0,
                FontSize = 10
            };

            LineSeries newSeries = new LineSeries
            {
                Title = "Register Monthly",
                Values = new ChartValues<int>(),
                DataLabels = true,
                LabelPoint = point => point.Y.ToString(),
                ScalesYAt = 0,
                FontSize = 10
            };
            EnumerableRowCollection<DataRow> filteredImportData = ImpData.AsEnumerable().Where(row => row.Field<DateTime>("Import_date").Year == currentYear);
            EnumerableRowCollection<DataRow> filteredExportData = ExpData.AsEnumerable().Where(row => row.Field<DateTime>("Export_date").Year == currentYear);
            EnumerableRowCollection<DataRow> filterednewData = DevData.AsEnumerable().Where(row => row.Field<DateTime>("Register_date").Year == currentYear);

            var importDataByMonth = filteredImportData.GroupBy(row => row.Field<DateTime>("Import_date").Month).Select(group => new
            {
                Month = group.Key,
                TotalImportQuantity = group.Sum(row => row.Field<int>("Import_quantity"))
            });

            var exportDataByMonth = filteredExportData.GroupBy(row => row.Field<DateTime>("Export_date").Month).Select(group => new
            {
                Month = group.Key,
                TotalExportQuantity = group.Sum(row => row.Field<int>("Export_quantity"))
            });

            var newDataByMonth = filterednewData.GroupBy(row => row.Field<DateTime>("Register_date").Month).Select(group => new
            {
                Month = group.Key,
                TotalExportQuantity = group.Sum(row => row.Field<int>("Quantity"))
            });

            for (int i = 1; i <= 12; i++)
            {
                int importQuantity = importDataByMonth.FirstOrDefault(item => item.Month == i)?.TotalImportQuantity ?? 0;
                int exportQuantity = exportDataByMonth.FirstOrDefault(item => item.Month == i)?.TotalExportQuantity ?? 0;
                int newQuantity = newDataByMonth.FirstOrDefault(item => item.Month == i)?.TotalExportQuantity ?? 0;

                importSeries.Values.Add(importQuantity);
                exportSeries.Values.Add(exportQuantity);
                newSeries.Values.Add(newQuantity);
            }

            imexp_chart.Series = seriesCollection;
            imexp_chart.Series.Add(importSeries);
            imexp_chart.Series.Add(exportSeries);
            imexp_chart.Series.Add(newSeries);
        }
    }
}