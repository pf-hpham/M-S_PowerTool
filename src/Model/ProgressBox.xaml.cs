using System.Windows;

namespace MnS
{
    public partial class ProgressBox : Window
    {
        public ProgressBox()
        {
            InitializeComponent();
        }

        public double ProgressValue
        {
            get { return progress.Value; }
            set { progress.Value = value; }
        }

        public void UpdateProgress(int value, string prog)
        {
            Dispatcher.Invoke(() => { progress.Value = value; });
            progbar.Content = $"Process running, please wait a moment... Loading {prog}%";
        }
    }
}