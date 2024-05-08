using MnS.ModelView;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;

namespace MnS
{
    public class CALModelView : INotifyPropertyChanged
    {
        private DataTable dt = new DataTable();

        public CALModelView(DataTable dataTable)
        {
            dt = dataTable;
            FillData(dt);
        }

        public ICommand RefreshCommand => new DelegateCommand(RefreshData);
        private ICollectionView collView;
        private string search;

        #region Observable
        public ObservableCollection<CALDevice> CALDevices { get; set; }
        public ObservableCollection<CALDevice> FilteredList { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        #endregion

        public string Search
        {
            get => search;
            set
            {
                search = value;

                collView.Filter = e =>
                {
                    var item = (CALDevice)e;
                    return item != null && ((item.Gage_ID?.StartsWith(search, StringComparison.OrdinalIgnoreCase) ?? false)
                                            || (item.Gage_SN?.StartsWith(search, StringComparison.OrdinalIgnoreCase) ?? false));
                };

                collView.Refresh();

                FilteredList = new ObservableCollection<CALDevice>(collView.OfType<CALDevice>().ToList());

                OnPropertyChanged("Search");
                OnPropertyChanged("FilteredList");
            }
        }

        private void FillData(DataTable dataTable)
        {
            try
            {
                if (dataTable == null)
                {
                    return;
                }

                search = "";

                CALDevices = new ObservableCollection<CALDevice>();

                foreach (DataRow row in dataTable.Rows)
                {
                    string nextDueDateString = row["Next_Due_Date"].ToString();
                    string lastCalibrationDateString = row["Last_Calibration_Date"].ToString();
                    DateTime nextDueDate;
                    DateTime lastCalibrationDate;

                    string dateFormat = "dd/MM/yyyy hh:mm:ss tt";

                    if (DateTime.TryParseExact(nextDueDateString, dateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out nextDueDate) &&
                        DateTime.TryParseExact(lastCalibrationDateString, dateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out lastCalibrationDate))
                    {
                        CALDevices.Add(new CALDevice
                        {
                            Gage_ID = row["Gage_ID"].ToString(),
                            Gage_SN = row["Gage_SN"].ToString(),
                            Model_No = row["Model_No"].ToString(),
                            Manufacturer = row["Manufacturer"].ToString(),
                            GM_Owner = row["GM_Owner"].ToString(),
                            Description = row["Description"].ToString(),
                            Current_Location = row["Current_Location"].ToString(),
                            Calibrator = row["Calibrator"].ToString(),
                            Calibration_Frequency = row["Calibration_Frequency"].ToString(),
                            Calibration_Frequency_UOM = row["Calibration_Frequency_UOM"].ToString(),
                            Next_Due_Date = nextDueDate,
                            Last_Calibration_Date = lastCalibrationDate,
                            Status = int.Parse(row["Status"].ToString()),
                            User_Defined = row["User_Defined"].ToString(),
                            Calibrated_By = row["Calibrated_By"].ToString()
                        });
                    }
                }

                FilteredList = new ObservableCollection<CALDevice>(CALDevices);
                collView = CollectionViewSource.GetDefaultView(FilteredList);

                OnPropertyChanged("Search");
                OnPropertyChanged("CALDevices");
                OnPropertyChanged("FilteredList");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex);
            }
        }

        private void OnPropertyChanged(string propertyname)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyname));
        }

        private void RefreshData(object obj)
        {
            FillData(dt);
        }
    }
}