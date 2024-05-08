using System;
using System.Windows;

namespace MnS
{
    public partial class Progress : Window
    {
        public static Progress instance;
        public string action;

        public Progress()
        {
            InitializeComponent();
            instance = this;
        }

        public void UpdateProgress(string action, int i)
        {
            progress.Value = i;
            actionList.Items.Add(action);
        }
    }
}