namespace UPDL_Speed_Tracker
{
    class Clear
    {
        //Reset Date & Time to Current
        public static void ResetDateTime(Form1 Form1)
        {
            GetCurrentDateTime.CurrentDate(Form1);
            GetCurrentDateTime.CurrentTime(Form1);
        }

        //Reset ComboBox & TextBox
        public static void ResetAll(Form1 Form1)
        {
            //Reset Cycle ComboBox to Default
            Form1.Cycle_ComboBox.SelectedIndex = 0;

            //Clear Items Count TextBox
            Form1.FileSize_TextBox.Clear();
        }
    }
}
