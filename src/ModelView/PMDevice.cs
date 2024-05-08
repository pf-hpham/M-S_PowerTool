using System;

namespace MnS
{
    public class PMDevice
    {
        /// <summary>
        /// 
        /// </summary>
        public PMDevice()
        {
        }
        
        #region PM Device
        /// <summary>
        /// 
        /// </summary>
        /// <param name="gageId"></param>
        /// <param name="gageSn"></param>
        /// <param name="assetNo"></param>
        /// <param name="modelNo"></param>
        /// <param name="manufacturer"></param>
        /// <param name="gmOwner"></param>
        /// <param name="description"></param>
        /// <param name="changeLevel"></param>
        /// <param name="storageLocation"></param>
        /// <param name="currentLocation"></param>
        /// <param name="calibrator"></param>
        /// <param name="calibrationFrequency"></param>
        /// <param name="calibrationFrequencyUom"></param>
        /// <param name="nextDueDate"></param>
        /// <param name="lastCalibrationDate"></param>
        /// <param name="status"></param>
        /// <param name="userDefined"></param>
        /// <param name="calibratedBy"></param>
        /// <param name="userDef1"></param>
        /// <param name="userDef2"></param>
        public PMDevice(string gageId, string gageSn, string assetNo, string modelNo, string manufacturer, string gmOwner, string description, string changeLevel, string storageLocation, string currentLocation, string calibrator, string calibrationFrequency, string calibrationFrequencyUom, DateTime nextDueDate, DateTime lastCalibrationDate, int status, string userDefined, string calibratedBy, string userDef1, string userDef2)
        {
            Gage_ID = gageId;
            Gage_SN = gageSn;
            Asset_No = assetNo;
            Model_No = modelNo;
            Manufacturer = manufacturer;
            GM_Owner = gmOwner;
            Description = description;
            Change_Level = changeLevel;
            Storage_Location = storageLocation;
            Current_Location = currentLocation;
            Calibrator = calibrator;
            Calibration_Frequency = calibrationFrequency;
            Calibration_Frequency_UOM = calibrationFrequencyUom;
            Next_Due_Date = nextDueDate;
            Last_Calibration_Date = lastCalibrationDate;
            Status = status;
            User_Defined = userDefined;
            Calibrated_By = calibratedBy;
            UserDef1 = userDef1;
            UserDef2 = userDef2;
        }
        public string Gage_ID { get; set; }
        public string Gage_SN { get; set; }
        public string Asset_No { get; set; }
        public string Model_No { get; set; }
        public string Manufacturer { get; set; }
        public string GM_Owner { get; set; }
        public string Description { get; set; }
        public string Change_Level { get; set; }
        public string Storage_Location { get; set; }
        public string Current_Location { get; set; }
        public string Calibrator { get; set; }
        public string Calibration_Frequency { get; set; }
        public string Calibration_Frequency_UOM { get; set; }
        public DateTime Next_Due_Date { get; set; }
        public DateTime Last_Calibration_Date { get; set; }
        public int Status { get; set; }
        public string User_Defined { get; set; }
        public string Calibrated_By { get; set; }
        public string UserDef1 { get; set; }
        public string UserDef2 { get; set; }
        #endregion
    }
}