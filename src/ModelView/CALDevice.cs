using System;

namespace MnS
{
    public class CALDevice
    {
        public CALDevice()
        {
        }

        #region Calibration Device
        public CALDevice(string gageId, string gageSn, string modelNo, string manufacturer, string gmOwner, string description, string currentLocation, string calibrator, string calibrationFrequency, string calibrationFrequencyUom, DateTime nextDueDate, DateTime lastCalibrationDate, int status, string userDefined, string calibratedBy)
        {
            Gage_ID = gageId;
            Gage_SN = gageSn;
            Model_No = modelNo;
            Manufacturer = manufacturer;
            GM_Owner = gmOwner;
            Description = description;
            Current_Location = currentLocation;
            Calibrator = calibrator;
            Calibration_Frequency = calibrationFrequency;
            Calibration_Frequency_UOM = calibrationFrequencyUom;
            Next_Due_Date = nextDueDate;
            Last_Calibration_Date = lastCalibrationDate;
            Status = status;
            User_Defined = userDefined;
            Calibrated_By = calibratedBy;
        }

        public string Gage_ID { get; set; }
        public string Gage_SN { get; set; }
        public string Model_No { get; set; }
        public string Manufacturer { get; set; }
        public string GM_Owner { get; set; }
        public string Description { get; set; }
        public string Current_Location { get; set; }
        public string Calibrator { get; set; }
        public string Calibration_Frequency { get; set; }
        public string Calibration_Frequency_UOM { get; set; }
        public DateTime Next_Due_Date { get; set; }
        public DateTime Last_Calibration_Date { get; set; }
        public int Status { get; set; }
        public string User_Defined { get; set; }
        public string Calibrated_By { get; set; }
        #endregion
    }
}
