using System;
using OxyPlot;
using OxyPlot.Axes;
using System.Windows;
using OxyPlot.Series;
using System.Collections.Generic;
using System.Data;
using OxyPlot.Wpf;
using Microsoft.Win32;
using OxyPlot.Legends;
using MnS.lib;

namespace MnS
{
    public partial class LineChart : Window
    {
        public static string LineChartTitle;
        private static PlotModel plotModel;

        public LineChart()
        {
            UserLogTool.UserData("Using Line Chart function");
            InitializeComponent();
            try
            {
                plotModel = new PlotModel();
                plotModel.Title = LineChartTitle;
                plotModel.Background = OxyColors.White;

                int n = 0;
                foreach (DataTable dt in BoxPlot.filter_tb)
                {
                    if (dt.Rows.Count > n)
                    {
                        n = dt.Rows.Count;
                    }
                }

                var xAxis = new LinearAxis
                {
                    Position = AxisPosition.Bottom,
                    MinimumPadding = 0.1,
                    MaximumPadding = 0.1,
                    AbsoluteMinimum = 0,
                    AbsoluteMaximum = n-1,
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot
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

                yAxis1.Labels.Add("Min: " + BoxPlot.min_v + BoxPlot.unit_v);
                yAxis1.Labels.Add("Median: " + BoxPlot.median + BoxPlot.unit_v);
                yAxis1.Labels.Add("Max: " + BoxPlot.max_v + BoxPlot.unit_v);

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

                plotModel.Series.Add(minLineSeries);
                plotModel.Series.Add(maxLineSeries);
                plotModel.Series.Add(medianLineSeries);

                List<DataTable> dataTableList = BoxPlot.filter_tb;

                for (int i = 0; i < dataTableList.Count; i++)
                {
                    var lineSeries = new LineSeries
                    {
                        Title = $"MO: {BoxPlot.mo_list[i]}",
                        StrokeThickness = 1.5,
                    };

                    for (int j = 0; j < dataTableList[i].Rows.Count; j++)
                    {
                        lineSeries.Points.Add(new DataPoint(j, Convert.ToDouble(dataTableList[i].Rows[j][0])));
                    }

                    plotModel.Series.Add(lineSeries);
                }

                for (int i = 0; i < n; i++)
                {
                    minLineSeries.Points.Add(new DataPoint(i, BoxPlot.min_v));
                    maxLineSeries.Points.Add(new DataPoint(i, BoxPlot.max_v));
                    medianLineSeries.Points.Add(new DataPoint(i, BoxPlot.median));
                }

                Legend legend = new Legend();
                legend.LegendPlacement = LegendPlacement.Outside;
                plotModel.Legends.Add(legend);

                LineChartPlotView.Model = plotModel;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex);
            }
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
                    PngExporter.Export(plotModel, saveFileDialog.FileName, 1220, 550, 96);

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