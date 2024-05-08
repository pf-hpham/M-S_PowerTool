using System;
using OxyPlot;
using OxyPlot.Axes;
using System.Windows;
using OxyPlot.Series;
using System.Collections.Generic;
using System.Data;
using OxyPlot.Wpf;
using Microsoft.Win32;
using MnS.lib;

namespace MnS
{
    public partial class BoxPlot : Window
    {
        #region Variable
        public static List<DataTable> filter_tb = new List<DataTable>();
        public static List<BoxPlotItem> list_box = new List<BoxPlotItem>();
        public static List<string> mo_list = new List<string>();
        public static List<string> date = new List<string>();
        private static PlotModel plotModel;
        public static string BoxPlotTitle;
        public static double min_v;
        public static double max_v;
        public static double median;
        public static string unit_v;
        public static int numberOfDataPoints;
        #endregion

        public BoxPlot()
        {
            UserLogTool.UserData("Using Box Plot function");
            InitializeComponent();
            #region BoxPlot
            try
            {
                plotModel = new PlotModel();
                plotModel.Title = BoxPlotTitle;
                plotModel.TitleHorizontalAlignment = TitleHorizontalAlignment.CenteredWithinView;
                plotModel.Background = OxyColors.White;
                numberOfDataPoints = filter_tb.Count;
                var boxPlotSeries = new BoxPlotSeries
                {
                    BoxWidth = 0.5,
                    WhiskerWidth = 0.5,
                    Fill = OxyColors.LightBlue
                };

                for (int i = 0; i < numberOfDataPoints; i++)
                {
                    boxPlotSeries.Items.Add(list_box[i]);
                }

                plotModel.Series.Add(boxPlotSeries);

                var xAxis = new CategoryAxis
                {
                    Position = AxisPosition.Bottom,
                    MinimumPadding = 0.1,
                    MaximumPadding = 0.1,
                    AbsoluteMinimum = -1,
                    AbsoluteMaximum = numberOfDataPoints,
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot,
                    Angle = - (numberOfDataPoints - 1)
                };

                var yAxis = new LinearAxis
                {
                    Position = AxisPosition.Left,
                    MinimumPadding = 0.1,
                    MaximumPadding = 0.1,
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot
                };

                var yAxis1 = new CategoryAxis
                {
                    Position = AxisPosition.Right,
                    MinimumPadding = 0.1,
                    MaximumPadding = 0.1,
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot,
                    Angle = 90
                };

                for (int i = 1; i <= numberOfDataPoints; i++)
                {
                    string label = mo_list[i - 1] + "\n" + date[i - 1];
                    xAxis.Labels.Add(label);
                }

                yAxis1.Labels.Add("Min: " + min_v + unit_v);
                yAxis1.Labels.Add("Median: " + median + unit_v);
                yAxis1.Labels.Add("Max: " + max_v + unit_v);

                plotModel.Axes.Add(xAxis);
                plotModel.Axes.Add(yAxis);
                plotModel.Axes.Add(yAxis1);

                var minLineSeries = new LineSeries
                {
                    Title = "Min Line",
                    LineStyle = LineStyle.Dash,
                    Color = OxyColors.Red
                };

                var maxLineSeries = new LineSeries
                {
                    Title = "Max Line",
                    LineStyle = LineStyle.Dash,
                    Color = OxyColors.Red
                };

                var medianLineSeries = new LineSeries
                {
                    Title = "Median Line",
                    LineStyle = LineStyle.Dash,
                    Color = OxyColors.Green
                };

                for (int i = -1; i <= numberOfDataPoints; i++)
                {
                    minLineSeries.Points.Add(new DataPoint(i, min_v));
                    maxLineSeries.Points.Add(new DataPoint(i, max_v));
                    medianLineSeries.Points.Add(new DataPoint(i, median));
                }

                plotModel.Series.Add(minLineSeries);
                plotModel.Series.Add(maxLineSeries);
                plotModel.Series.Add(medianLineSeries);

                RegressionDiagram.Model = plotModel;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex);
            }
            #endregion
        }

        private void SaveChartAsImage(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Title = "Save Chart As Image",
                    Filter = "PNG Files (*.png)|*.png|All Files (*.*)|*.*",
                    DefaultExt = "png"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    PngExporter.Export(plotModel, saveFileDialog.FileName, (int)RegressionDiagram.ActualWidth, (int)RegressionDiagram.ActualHeight, 96);

                    MessageBox.Show($"Chart saved as: {saveFileDialog.FileName}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving chart: {ex.Message}");
            }
        }
    }
}