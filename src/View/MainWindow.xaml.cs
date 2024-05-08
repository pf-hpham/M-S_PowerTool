using MnS.lib;
using System;
using System.Windows;
using System.Reflection;


namespace MnS
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DisplayAppVersion();
        }

        private void DisplayAppVersion()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Version version = assembly.GetName().Version;
            App_version.Content = $"Version {version}";
        }

        public void Image_Viewer(object sender, RoutedEventArgs e)
        {
            ImageViewer imageViewer = new ImageViewer();
            imageViewer.Image_Click(sender, e);
        }
    }
}