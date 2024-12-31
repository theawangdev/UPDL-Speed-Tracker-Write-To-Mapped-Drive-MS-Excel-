using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace UPDL_Speed_Tracker
{
    class Submit
    {
        private static string ExcelPath = GetAppSetting.Get("ExcelPath"); //Location of Excel file
        private static string SpreadsheetName = GetAppSetting.Get("SpreadsheetName"); //Location to write data in Excel
        private static int startRow = 7; //Start write data from row number 6
        private static int nextRow;

        private static ManualResetEvent _pauseEvent = new ManualResetEvent(false);

        public static void SubmitData(Form1 Form1)
        {
            // Collect data from form fields
            string chooseDateSubmit = Form1.SubmitDate_DatePicker.Value.ToString("dd/MM/yyyy");

            string uploadStartTime = Form1.UPLOADStart_TimePicker.Value.ToString("T");
            string uploadEndTime = Form1.UPLOADEnd_TimePicker.Value.ToString("T");

            string downloadStartTime = Form1.DOWNLOADStart_TimePicker.Value.ToString("T");
            string downloadEndTime = Form1.DOWNLOADEnd_TimePicker.Value.ToString("T");

            int fileSize = int.Parse(Form1.FileSize_TextBox.Text);

            WriteToExcel(chooseDateSubmit, uploadStartTime, uploadEndTime, downloadStartTime, downloadEndTime, fileSize, Form1);

            MessageBox.Show("Data submitted successfully!");
        }

        public static void WriteToExcel(String chooseDateSubmit, string uploadStartTime, string uploadEndTime, string downloadStartTime,
            string downloadEndTime, int fileSize, Form1 Form1)
        {
            Excel.Application excelApp = new Excel.Application();

            Excel.Workbook workbook = excelApp.Workbooks.Open(ExcelPath);
            Excel.Worksheet worksheet = workbook.Sheets[SpreadsheetName];
            excelApp.Visible = false;

            try
            {
                nextRow = startRow;

                switch (Form1.Cycle_ComboBox.SelectedItem)
                {
                    case "Cycle 1":

                        //Loop to find the next available row
                        while (true)
                        {
                            //Check cells in the row for data
                            Excel.Range dateCycleCell = worksheet.Cells[nextRow, 2]; // Column B
                            Excel.Range uploadCell = worksheet.Cells[nextRow, 4]; // Column D
                            Excel.Range downloadCell = worksheet.Cells[nextRow, 7]; // Column G
                            Excel.Range fileSizeCell = worksheet.Cells[nextRow, 10]; // Column J

                            if (dateCycleCell.Value == null && uploadCell.Value == null && downloadCell.Value == null && fileSizeCell.Value == null)
                            {
                                //If the row has no data, break the loop to write data
                                break;
                            }

                            //Move to the next row if row has data
                            nextRow++;
                        }

                        //Write data into next available row
                        //Number is refer to Column Alphabet
                        worksheet.Cells[nextRow, 2].Value = chooseDateSubmit; //Column B
                        worksheet.Cells[nextRow, 4].Value = uploadStartTime; //Column D
                        worksheet.Cells[nextRow, 5].Value = uploadEndTime; //Column E
                        worksheet.Cells[nextRow, 7].Value = downloadStartTime; //Column G
                        worksheet.Cells[nextRow, 8].Value = downloadEndTime; //Column H
                        worksheet.Cells[nextRow, 10].Value = fileSize; //Column J
                        break;

                    case "Cycle 2":

                        //Loop to find the next available row
                        while (true)
                        {
                            //Check cells in the row for data
                            Excel.Range uploadCell = worksheet.Cells[nextRow, 16]; //Column P
                            Excel.Range downloadCell = worksheet.Cells[nextRow, 19]; //Column S
                            Excel.Range fileSizeCell = worksheet.Cells[nextRow, 22]; //Column V

                            if (uploadCell.Value == null && downloadCell.Value == null && fileSizeCell.Value == null)
                            {
                                //If the row has no data, break the loop
                                break;
                            }

                            //Move to the next row
                            nextRow++;
                        }

                        //Set the data into the next available row
                        worksheet.Cells[nextRow, 16].Value = uploadStartTime; //Column P
                        worksheet.Cells[nextRow, 17].Value = uploadEndTime; //Column Q
                        worksheet.Cells[nextRow, 19].Value = downloadStartTime; //Column S
                        worksheet.Cells[nextRow, 20].Value = downloadEndTime; //Column T
                        worksheet.Cells[nextRow, 22].Value = fileSize; //Column V
                        break;

                    case "Cycle 3":

                        //Loop to find the next available row
                        while (true)
                        {
                            //Check cells in the row for data
                            Excel.Range uploadCell = worksheet.Cells[nextRow, 28]; //Column AB
                            Excel.Range downloadCell = worksheet.Cells[nextRow, 31]; //Column AE
                            Excel.Range fileSizeCell = worksheet.Cells[nextRow, 34]; //Column AH

                            if (uploadCell.Value == null && downloadCell.Value == null && fileSizeCell.Value == null)
                            {
                                //If the row has no data, break the loop
                                break;
                            }

                            //Move to the next row
                            nextRow++;
                        }

                        //Set the data into the next available row
                        worksheet.Cells[nextRow, 28].Value = uploadStartTime; //Column AB
                        worksheet.Cells[nextRow, 29].Value = uploadEndTime; //Column AC
                        worksheet.Cells[nextRow, 31].Value = downloadStartTime; //Column AE
                        worksheet.Cells[nextRow, 32].Value = downloadEndTime; //Column AF
                        worksheet.Cells[nextRow, 34].Value = fileSize; //Column AH
                        break;

                    case "Cycle 4":

                        //Loop to find the next available row
                        while (true)
                        {
                            //Check cells in the row for data
                            Excel.Range uploadCell = worksheet.Cells[nextRow, 40]; //Column AN
                            Excel.Range downloadCell = worksheet.Cells[nextRow, 43]; //Column AQ
                            Excel.Range fileSizeCell = worksheet.Cells[nextRow, 46]; //Column AT

                            if (uploadCell.Value == null && downloadCell.Value == null && fileSizeCell.Value == null)
                            {
                                //If the row has no data, break the loop
                                break;
                            }

                            //Move to the next row
                            nextRow++;
                        }

                        //Set the data into the next available row
                        worksheet.Cells[nextRow, 40].Value = uploadStartTime; //Column AN
                        worksheet.Cells[nextRow, 41].Value = uploadEndTime; //Column AO
                        worksheet.Cells[nextRow, 43].Value = downloadStartTime; //Column AQ
                        worksheet.Cells[nextRow, 44].Value = downloadEndTime; //Column AR
                        worksheet.Cells[nextRow, 46].Value = fileSize; //Column AT
                        break;
                }
                
                //Save the workbook
                workbook.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

            //Close the workbook & release Excel process like unbind
            workbook.Close();
            releaseObject(workbook);
            excelApp.Quit();
            releaseObject(excelApp);
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
