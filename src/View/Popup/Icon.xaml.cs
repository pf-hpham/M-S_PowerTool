using MnS.lib;
using System.Windows;

namespace MnS
{
    public partial class Icon : Window
    {
        public Icon()
        {
            UserLogTool.UserData("Using Icon function");
            InitializeComponent();
            Top = SystemParameters.WorkArea.Height - Height;
            Left = SystemParameters.WorkArea.Width / 2;
        }

        private void Ic_Click(object sender, RoutedEventArgs e)
        {
            MiniTool mini = new MiniTool();
            mini.Show();
            Hide();
        }
    }
}
