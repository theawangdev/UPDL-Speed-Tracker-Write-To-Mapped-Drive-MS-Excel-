using System;

namespace UPDL_Speed_Tracker
{
    class GetCurrentDateTime
    {
        //Get current Date
        public static void CurrentDate(Form1 Form1)
        {
            Form1.SubmitDate_DatePicker.Value = DateTime.Now.Date;
        }
        
        //
        //Get current Time
        public static void CurrentTime(Form1 Form1)
        {
            Form1.UPLOADStart_TimePicker.Value = DateTime.Now;
            Form1.UPLOADEnd_TimePicker.Value = DateTime.Now;

            Form1.DOWNLOADStart_TimePicker.Value = DateTime.Now;
            Form1.DOWNLOADEnd_TimePicker.Value = DateTime.Now;
        }
    }
}
