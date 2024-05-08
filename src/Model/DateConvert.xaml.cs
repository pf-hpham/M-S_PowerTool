using MnS.lib;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace MnS
{
    public partial class DateConvert : Window
    {
        public DateTime LastMonthDate { get; set; }
        public DateTime ThisMonthDate { get; set; }
        public DateTime NextMonthDate { get; set; }

        public DateConvert()
        {
            UserLogTool.UserData("Using Date Converter function");
            InitializeComponent(); 
            SetCalendarDates();
            DataContext = this;
            datePicker.SelectedDateChanged += DatePicker_SelectedDateChanged;
        }

        private void SetCalendarDates()
        {
            DateTime currentDate = DateTime.Today;
            LastMonthDate = currentDate.AddMonths(-1);
            ThisMonthDate = currentDate;
            NextMonthDate = currentDate.AddMonths(1);
        }

        private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (datePicker.SelectedDate.HasValue)
            {
                DateTime selectedDate = datePicker.SelectedDate.Value;

                CultureInfo cultureInfo = new CultureInfo("en-US");
                int weekOfYear = cultureInfo.Calendar.GetWeekOfYear(selectedDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

                int Year = cultureInfo.Calendar.GetYear(selectedDate);

                resultTextBlock.Text = $"Week {weekOfYear} of {Year}";
            }
            else
            {
                resultTextBlock.Text = "Please select a date.";
            }
        }

    }
}