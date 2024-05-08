using LiveCharts;
using LiveCharts.Defaults;
using LiveCharts.Wpf;
using System;
using System.Data;
using System.Windows;
using System.Windows.Media;

namespace MnS
{
    public partial class DAT_Graph : Window
    {
        #region
        double lower;
        double upper;
        string columnName;
        DataTable data;

        CartesianChart chart;
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Lower"></param>
        /// <param name="Upper"></param>
        /// <param name="Data"></param>
        /// <param name="column_Name"></param>
        public DAT_Graph(double Lower, double Upper, DataTable Data, string column_Name)
        {
            InitializeComponent();
            lower = Lower;
            upper = Upper;
            columnName = column_Name;
            foreach (DataRow row in Data.Rows)
            {
                string cellValue = row[columnName].ToString();
                if (cellValue.Contains("_OK"))
                {
                    row[columnName] = cellValue.Replace("_OK", "");
                }
            }

            data = Data;
            Graph_Build();
            Grid.Children.Add(chart);
        }

        /// <summary>
        /// 
        /// </summary>
        public void Graph_Build()
        {
            chart = new CartesianChart();

            // Tạo Series cho upper line
            var upperLineSeries = new LineSeries
            {
                Title = "Upper Line",
                Values = new ChartValues<ObservablePoint>(),
                Stroke = Brushes.Red,
                Fill = Brushes.Transparent
            };

            // Tạo Series cho lower line
            var lowerLineSeries = new LineSeries
            {
                Title = "Lower Line",
                Values = new ChartValues<ObservablePoint>(),
                Stroke = Brushes.Blue,
                Fill = Brushes.Transparent
            };

            // Tạo Series cho average line
            var averageLineSeries = new LineSeries
            {
                Title = "Average Line",
                Values = new ChartValues<ObservablePoint>(),
                Stroke = Brushes.Green,
                Fill = Brushes.Transparent
            };

            // Tạo Series cho giá trị thay đổi
            var valueChangeSeries = new LineSeries
            {
                Title = "Value Change",
                Values = new ChartValues<double>(),
                Stroke = Brushes.Orange,
                Fill = Brushes.Transparent
            };

            for (int i = 0; i < data.Rows.Count; i++)
            {
                double value = Convert.ToDouble(data.Rows[i][columnName]);
                double average = (upper + lower) / 2;

                // Thêm điểm cho upper line
                upperLineSeries.Values.Add(new ObservablePoint(i, upper));

                // Thêm điểm cho lower line
                lowerLineSeries.Values.Add(new ObservablePoint(i, lower));

                // Thêm điểm cho average line
                averageLineSeries.Values.Add(new ObservablePoint(i, average));

                // Thêm giá trị cho giá trị thay đổi
                valueChangeSeries.Values.Add(value);
            }

            // Thêm các series vào biểu đồ
            chart.Series = new SeriesCollection
            {
                upperLineSeries,
                lowerLineSeries,
                averageLineSeries,
                valueChangeSeries
            };

            // Cài đặt trục X và Y
            chart.AxisX.Add(new Axis());
            chart.AxisY.Add(new Axis());
        }
    }
}