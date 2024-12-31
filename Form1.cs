using System;
using System.Drawing;
using System.Windows.Forms;

namespace UPDL_Speed_Tracker
{
    public partial class Form1 : Form
    {
        //Top Panel Variable
        private bool isDragging = false;
        private Point cursorStartPoint;

        public Form1()
        {
            InitializeComponent();

            //Reset Cycle ComboBox to Default
            Cycle_ComboBox.SelectedIndex = 0;

            //Show current Date & Time
            GetCurrentDateTime.CurrentDate(this);
            GetCurrentDateTime.CurrentTime(this);
        }

        //
        //Top Panel: Set to able move Form
        private void TopPanel_MouseDown(object sender, MouseEventArgs e)
        {
            isDragging = true;

            cursorStartPoint = new Point(e.X, e.Y);
        }

        private void TopPanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Point pointScreen = PointToScreen(e.Location);
                this.Location = new Point(pointScreen.X - cursorStartPoint.X, pointScreen.Y - cursorStartPoint.Y);
            }
        }

        private void TopPanel_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
        }

        //
        //Minimize Label: Minimize Application
        private void MinimizeLabel_Click(object sender, System.EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        //
        //Close Label: Close Application
        private void CloseLabel_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        //
        //Bottom Panel: Border Line
        private void BottomPanel_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, BottomPanel.ClientRectangle,
                Color.DarkCyan, 1, ButtonBorderStyle.Solid,
                Color.Black, 0, ButtonBorderStyle.None,
                Color.DarkCyan, 1, ButtonBorderStyle.Solid,
                Color.DarkCyan, 1, ButtonBorderStyle.Solid);
        }

        //
        //Items Count TextBox: Set only can input number
        private void ItemsCount_TextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar))
            {
                if (!char.IsDigit(e.KeyChar))
                {
                    e.Handled = true;
                }
            }
        }

        //
        //Submit Button: Store data to Google Sheet
        private void SubmitButton_Click(object sender, EventArgs e)
        {
            //Get user input and write data into Excel
            Submit.SubmitData(this);

            ////Reset All field to Default after success submit data
            Clear.ResetDateTime(this);
            Clear.ResetAll(this);
        }
    }
}
