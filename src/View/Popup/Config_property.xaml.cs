using System;
using System.IO;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using MnS.lib;

namespace MnS
{
    public class ConfigItem
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public partial class Config_property : Window
    {
        List<ConfigItem> dataItems = new List<ConfigItem>();

        public string Path { get; set; }

        public Config_property(string filepath)
        {
            UserLogTool.UserData("Open config function");
            InitializeComponent();
            Path = filepath;
            LoadTextFile(filepath);
        }

        private void LoadTextFile(string filePath)
        {
            try
            {
                string[] lines = File.ReadAllLines(filePath);

                foreach (var line in lines)
                {
                    string[] parts = line.Split(new char[] { '=' }, 2);

                    if (parts.Length == 2)
                    {
                        ConfigItem item = new ConfigItem
                        {
                            Key = parts[0].Trim(),
                            Value = AddNewLineBeforeKeywords(parts[1].Trim())
                        };

                        dataItems.Add(item);
                    }
                    else
                    {
                        ConfigItem item = new ConfigItem
                        {
                            Key = line.Trim(),
                            Value = string.Empty
                        };

                        dataItems.Add(item);
                    }
                }

                config_grid.ItemsSource = dataItems;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private string AddNewLineBeforeKeywords(string value)
        {
            string[] keywords = { "from", "where", "order by" };

            foreach (var keyword in keywords)
            {
                value = value.Replace(keyword, Environment.NewLine + keyword);
            }

            return value;
        }

        private void ConfigGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            int rowIndex = e.Row.GetIndex();
            int valueColumnIndex = 1;

            if (e.Column.DisplayIndex == valueColumnIndex && rowIndex >= 0)
            {
                var editedValue = ((TextBox)e.EditingElement).Text;
                dataItems[rowIndex].Value = AddNewLineBeforeKeywords(editedValue);
            }
        }

        private void SaveChangesToFile(string filePath)
        {
            try
            {
                List<string> linesToWrite = new List<string>();

                foreach (var item in dataItems)
                {
                    if(item.Value != "")
                    {
                        string value = RemoveNewLineBeforeKeywords(item.Value);
                        string line = $"{item.Key}={value}";
                        linesToWrite.Add(line);
                    }
                    else
                    {
                        string line = $"{item.Key}";
                        linesToWrite.Add(line);
                    }
                }

                File.WriteAllLines(filePath, linesToWrite);
                MessageBox.Show("Save file successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private string RemoveNewLineBeforeKeywords(string value)
        {
            string[] keywords = { "from", "where", "order by" };

            foreach (var keyword in keywords)
            {
                value = value.Replace(Environment.NewLine + keyword, keyword);
            }

            return value;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            SaveChangesToFile(Path);
        }
    }
}
