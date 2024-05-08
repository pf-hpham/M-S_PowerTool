using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using System.Windows.Input;

namespace MnS.ModelView
{
    public class PMModelView : INotifyPropertyChanged
    {
        private DataTable dt = new DataTable();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataTable"></param>
        public PMModelView(DataTable dataTable)
        {
            dt = dataTable;
            FillData(dt);
        }

        /// <summary>
        ///     Refresh all
        /// </summary>
        public ICommand RefreshCommand => new DelegateCommand(RefreshData);
        private ICollectionView collView;
        private string search;

        #region Observable
        public ObservableCollection<PMDevice> PMDevices { get; set; }
        public ObservableCollection<PMDevice> FilteredList { get; set; }
        private void OnPropertyChanged(string propertyname)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyname));
        }

        public event PropertyChangedEventHandler PropertyChanged;
        #endregion

        /// <summary>
        /// Global filter
        /// </summary>
        public string Search
        {
            get => search;
            set
            {
                search = value;

                collView.Filter = e =>
                {
                    var item = (PMDevice)e;
                    return item != null && ((item.Gage_ID?.StartsWith(search, StringComparison.OrdinalIgnoreCase) ?? false)
                                            || (item.Gage_SN?.StartsWith(search, StringComparison.OrdinalIgnoreCase) ?? false));
                };

                collView.Refresh();

                FilteredList = new ObservableCollection<PMDevice>(collView.OfType<PMDevice>().ToList());

                OnPropertyChanged("Search");
                OnPropertyChanged("FilteredList");
            }
        }

        /// <summary>
        /// Fill data
        /// </summary>
        private void FillData(DataTable dataTable)
        {
            if (dataTable == null)
            {
                return;
            }

            search = "";

            PMDevices = new ObservableCollection<PMDevice>();

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
                    PMDevices.Add(new PMDevice
                    {
                        Gage_ID = row["Gage_ID"].ToString(),
                        Gage_SN = row["Gage_SN"].ToString(),
                        Asset_No = row["Asset_No"].ToString(),
                        Model_No = row["Model_No"].ToString(),
                        Manufacturer = row["Manufacturer"].ToString(),
                        GM_Owner = row["GM_Owner"].ToString(),
                        Description = row["Description"].ToString(),
                        Change_Level = row["Change_Level"].ToString(),
                        Storage_Location = row["Storage_Location"].ToString(),
                        Current_Location = row["Current_Location"].ToString(),
                        Calibrator = row["Calibrator"].ToString(),
                        Calibration_Frequency = row["Calibration_Frequency"].ToString(),
                        Calibration_Frequency_UOM = row["Calibration_Frequency_UOM"].ToString(),
                        Next_Due_Date = DateTime.Parse(row["Next_Due_Date"].ToString()), // Cần phải phân tích chuỗi thành DateTime
                        Last_Calibration_Date = DateTime.Parse(row["Last_Calibration_Date"].ToString()), // Cần phải phân tích chuỗi thành DateTime
                        Status = int.Parse(row["Status"].ToString()),
                        User_Defined = row["User_Defined"].ToString(),
                        Calibrated_By = row["Calibrated_By"].ToString(),
                        UserDef1 = row["UserDef1"].ToString(),
                        UserDef2 = row["UserDef2"].ToString()
                    });
                }
                else
                {
                }
            }

            FilteredList = new ObservableCollection<PMDevice>(PMDevices);
            collView = CollectionViewSource.GetDefaultView(FilteredList);

            OnPropertyChanged("Search");
            OnPropertyChanged("PMDevices");
            OnPropertyChanged("FilteredList");
        }

        /// <summary>
        /// refresh data
        /// </summary>
        /// <param name="obj"></param>
        private void RefreshData(object obj)
        {
            FillData(dt);
        }
    }
}