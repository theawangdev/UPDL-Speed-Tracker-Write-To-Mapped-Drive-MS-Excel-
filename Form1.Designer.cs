
namespace UPDL_Speed_Tracker
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.MainPanel = new System.Windows.Forms.Panel();
            this.BottomPanel = new System.Windows.Forms.Panel();
            this.DOWNLOAD_GroupBox = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.DOWNLOADStart_TimePicker = new System.Windows.Forms.DateTimePicker();
            this.DOWNLOADEnd_TimePicker = new System.Windows.Forms.DateTimePicker();
            this.UPLOAD_GroupBox = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.UPLOADStart_TimePicker = new System.Windows.Forms.DateTimePicker();
            this.UPLOADEnd_TimePicker = new System.Windows.Forms.DateTimePicker();
            this.Cycle_ComboBox = new System.Windows.Forms.ComboBox();
            this.FileSize_TextBox = new System.Windows.Forms.TextBox();
            this.SubmitButton = new System.Windows.Forms.Button();
            this.SubmitDate_DatePicker = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.TopPanel = new System.Windows.Forms.Panel();
            this.label8 = new System.Windows.Forms.Label();
            this.FormControl_Panel = new System.Windows.Forms.Panel();
            this.MinimizeLabel = new System.Windows.Forms.Label();
            this.CloseLabel = new System.Windows.Forms.Label();
            this.ApplicationName_Label = new System.Windows.Forms.Label();
            this.MainPanel.SuspendLayout();
            this.BottomPanel.SuspendLayout();
            this.DOWNLOAD_GroupBox.SuspendLayout();
            this.UPLOAD_GroupBox.SuspendLayout();
            this.TopPanel.SuspendLayout();
            this.FormControl_Panel.SuspendLayout();
            this.SuspendLayout();
            // 
            // MainPanel
            // 
            this.MainPanel.Controls.Add(this.BottomPanel);
            this.MainPanel.Controls.Add(this.TopPanel);
            this.MainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.MainPanel.Location = new System.Drawing.Point(0, 0);
            this.MainPanel.Name = "MainPanel";
            this.MainPanel.Size = new System.Drawing.Size(544, 250);
            this.MainPanel.TabIndex = 0;
            // 
            // BottomPanel
            // 
            this.BottomPanel.BackColor = System.Drawing.Color.White;
            this.BottomPanel.Controls.Add(this.DOWNLOAD_GroupBox);
            this.BottomPanel.Controls.Add(this.UPLOAD_GroupBox);
            this.BottomPanel.Controls.Add(this.Cycle_ComboBox);
            this.BottomPanel.Controls.Add(this.FileSize_TextBox);
            this.BottomPanel.Controls.Add(this.SubmitButton);
            this.BottomPanel.Controls.Add(this.SubmitDate_DatePicker);
            this.BottomPanel.Controls.Add(this.label3);
            this.BottomPanel.Controls.Add(this.label2);
            this.BottomPanel.Controls.Add(this.label1);
            this.BottomPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.BottomPanel.Location = new System.Drawing.Point(0, 40);
            this.BottomPanel.Name = "BottomPanel";
            this.BottomPanel.Size = new System.Drawing.Size(544, 210);
            this.BottomPanel.TabIndex = 1;
            this.BottomPanel.Paint += new System.Windows.Forms.PaintEventHandler(this.BottomPanel_Paint);
            // 
            // DOWNLOAD_GroupBox
            // 
            this.DOWNLOAD_GroupBox.Controls.Add(this.label6);
            this.DOWNLOAD_GroupBox.Controls.Add(this.label7);
            this.DOWNLOAD_GroupBox.Controls.Add(this.DOWNLOADStart_TimePicker);
            this.DOWNLOAD_GroupBox.Controls.Add(this.DOWNLOADEnd_TimePicker);
            this.DOWNLOAD_GroupBox.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DOWNLOAD_GroupBox.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(29)))), ((int)(((byte)(29)))), ((int)(((byte)(29)))));
            this.DOWNLOAD_GroupBox.Location = new System.Drawing.Point(325, 62);
            this.DOWNLOAD_GroupBox.Name = "DOWNLOAD_GroupBox";
            this.DOWNLOAD_GroupBox.Size = new System.Drawing.Size(204, 94);
            this.DOWNLOAD_GroupBox.TabIndex = 16;
            this.DOWNLOAD_GroupBox.TabStop = false;
            this.DOWNLOAD_GroupBox.Text = "Download File";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(29)))), ((int)(((byte)(29)))), ((int)(((byte)(29)))));
            this.label6.Location = new System.Drawing.Point(11, 29);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(70, 17);
            this.label6.TabIndex = 3;
            this.label6.Text = "Start Time:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(29)))), ((int)(((byte)(29)))), ((int)(((byte)(29)))));
            this.label7.Location = new System.Drawing.Point(11, 60);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 17);
            this.label7.TabIndex = 4;
            this.label7.Text = "End Time:";
            // 
            // DOWNLOADStart_TimePicker
            // 
            this.DOWNLOADStart_TimePicker.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DOWNLOADStart_TimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.DOWNLOADStart_TimePicker.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.DOWNLOADStart_TimePicker.Location = new System.Drawing.Point(82, 26);
            this.DOWNLOADStart_TimePicker.Name = "DOWNLOADStart_TimePicker";
            this.DOWNLOADStart_TimePicker.ShowUpDown = true;
            this.DOWNLOADStart_TimePicker.Size = new System.Drawing.Size(106, 25);
            this.DOWNLOADStart_TimePicker.TabIndex = 8;
            this.DOWNLOADStart_TimePicker.Value = new System.DateTime(2024, 5, 27, 18, 5, 10, 0);
            // 
            // DOWNLOADEnd_TimePicker
            // 
            this.DOWNLOADEnd_TimePicker.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DOWNLOADEnd_TimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.DOWNLOADEnd_TimePicker.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.DOWNLOADEnd_TimePicker.Location = new System.Drawing.Point(82, 57);
            this.DOWNLOADEnd_TimePicker.Name = "DOWNLOADEnd_TimePicker";
            this.DOWNLOADEnd_TimePicker.ShowUpDown = true;
            this.DOWNLOADEnd_TimePicker.Size = new System.Drawing.Size(106, 25);
            this.DOWNLOADEnd_TimePicker.TabIndex = 9;
            this.DOWNLOADEnd_TimePicker.Value = new System.DateTime(2024, 5, 27, 18, 5, 10, 0);
            // 
            // UPLOAD_GroupBox
            // 
            this.UPLOAD_GroupBox.Controls.Add(this.label4);
            this.UPLOAD_GroupBox.Controls.Add(this.label5);
            this.UPLOAD_GroupBox.Controls.Add(this.UPLOADStart_TimePicker);
            this.UPLOAD_GroupBox.Controls.Add(this.UPLOADEnd_TimePicker);
            this.UPLOAD_GroupBox.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UPLOAD_GroupBox.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(29)))), ((int)(((byte)(29)))), ((int)(((byte)(29)))));
            this.UPLOAD_GroupBox.Location = new System.Drawing.Point(15, 62);
            this.UPLOAD_GroupBox.Name = "UPLOAD_GroupBox";
            this.UPLOAD_GroupBox.Size = new System.Drawing.Size(204, 94);
            this.UPLOAD_GroupBox.TabIndex = 15;
            this.UPLOAD_GroupBox.TabStop = false;
            this.UPLOAD_GroupBox.Text = "Upload File";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(29)))), ((int)(((byte)(29)))), ((int)(((byte)(29)))));
            this.label4.Location = new System.Drawing.Point(11, 29);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(70, 17);
            this.label4.TabIndex = 3;
            this.label4.Text = "Start Time:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(29)))), ((int)(((byte)(29)))), ((int)(((byte)(29)))));
            this.label5.Location = new System.Drawing.Point(11, 60);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 17);
            this.label5.TabIndex = 4;
            this.label5.Text = "End Time:";
            // 
            // UPLOADStart_TimePicker
            // 
            this.UPLOADStart_TimePicker.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UPLOADStart_TimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.UPLOADStart_TimePicker.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.UPLOADStart_TimePicker.Location = new System.Drawing.Point(82, 26);
            this.UPLOADStart_TimePicker.Name = "UPLOADStart_TimePicker";
            this.UPLOADStart_TimePicker.ShowUpDown = true;
            this.UPLOADStart_TimePicker.Size = new System.Drawing.Size(106, 25);
            this.UPLOADStart_TimePicker.TabIndex = 8;
            this.UPLOADStart_TimePicker.Value = new System.DateTime(2024, 5, 27, 18, 5, 10, 0);
            // 
            // UPLOADEnd_TimePicker
            // 
            this.UPLOADEnd_TimePicker.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.UPLOADEnd_TimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.UPLOADEnd_TimePicker.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.UPLOADEnd_TimePicker.Location = new System.Drawing.Point(82, 57);
            this.UPLOADEnd_TimePicker.Name = "UPLOADEnd_TimePicker";
            this.UPLOADEnd_TimePicker.ShowUpDown = true;
            this.UPLOADEnd_TimePicker.Size = new System.Drawing.Size(106, 25);
            this.UPLOADEnd_TimePicker.TabIndex = 9;
            this.UPLOADEnd_TimePicker.Value = new System.DateTime(2024, 5, 27, 18, 5, 10, 0);
            // 
            // Cycle_ComboBox
            // 
            this.Cycle_ComboBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(244)))), ((int)(((byte)(200)))));
            this.Cycle_ComboBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Cycle_ComboBox.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Cycle_ComboBox.FormattingEnabled = true;
            this.Cycle_ComboBox.Items.AddRange(new object[] {
            "Cycle 1",
            "Cycle 2",
            "Cycle 3",
            "Cycle 4"});
            this.Cycle_ComboBox.Location = new System.Drawing.Point(208, 19);
            this.Cycle_ComboBox.Name = "Cycle_ComboBox";
            this.Cycle_ComboBox.Size = new System.Drawing.Size(90, 25);
            this.Cycle_ComboBox.Sorted = true;
            this.Cycle_ComboBox.TabIndex = 14;
            // 
            // FileSize_TextBox
            // 
            this.FileSize_TextBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(244)))), ((int)(((byte)(200)))));
            this.FileSize_TextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.FileSize_TextBox.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FileSize_TextBox.Location = new System.Drawing.Point(384, 19);
            this.FileSize_TextBox.MaximumSize = new System.Drawing.Size(90, 25);
            this.FileSize_TextBox.MinimumSize = new System.Drawing.Size(90, 25);
            this.FileSize_TextBox.Name = "FileSize_TextBox";
            this.FileSize_TextBox.Size = new System.Drawing.Size(90, 25);
            this.FileSize_TextBox.TabIndex = 13;
            this.FileSize_TextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ItemsCount_TextBox_KeyPress);
            // 
            // SubmitButton
            // 
            this.SubmitButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(10)))), ((int)(((byte)(198)))), ((int)(((byte)(133)))));
            this.SubmitButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.SubmitButton.FlatAppearance.BorderSize = 0;
            this.SubmitButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(155)))), ((int)(((byte)(102)))));
            this.SubmitButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(10)))), ((int)(((byte)(198)))), ((int)(((byte)(133)))));
            this.SubmitButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SubmitButton.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SubmitButton.ForeColor = System.Drawing.Color.White;
            this.SubmitButton.Location = new System.Drawing.Point(15, 166);
            this.SubmitButton.Name = "SubmitButton";
            this.SubmitButton.Size = new System.Drawing.Size(514, 33);
            this.SubmitButton.TabIndex = 12;
            this.SubmitButton.Text = "SUBMIT";
            this.SubmitButton.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.SubmitButton.UseVisualStyleBackColor = false;
            this.SubmitButton.Click += new System.EventHandler(this.SubmitButton_Click);
            // 
            // SubmitDate_DatePicker
            // 
            this.SubmitDate_DatePicker.CalendarMonthBackground = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(244)))), ((int)(((byte)(200)))));
            this.SubmitDate_DatePicker.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SubmitDate_DatePicker.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.SubmitDate_DatePicker.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.SubmitDate_DatePicker.Location = new System.Drawing.Point(56, 19);
            this.SubmitDate_DatePicker.Name = "SubmitDate_DatePicker";
            this.SubmitDate_DatePicker.Size = new System.Drawing.Size(93, 25);
            this.SubmitDate_DatePicker.TabIndex = 7;
            this.SubmitDate_DatePicker.Value = new System.DateTime(2024, 5, 1, 0, 0, 0, 0);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(29)))), ((int)(((byte)(29)))), ((int)(((byte)(29)))));
            this.label3.Location = new System.Drawing.Point(321, 22);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 17);
            this.label3.TabIndex = 2;
            this.label3.Text = "File Size:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(29)))), ((int)(((byte)(29)))), ((int)(((byte)(29)))));
            this.label2.Location = new System.Drawing.Point(161, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 17);
            this.label2.TabIndex = 1;
            this.label2.Text = "Cycle:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(29)))), ((int)(((byte)(29)))), ((int)(((byte)(29)))));
            this.label1.Location = new System.Drawing.Point(12, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Date:";
            // 
            // TopPanel
            // 
            this.TopPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(155)))), ((int)(((byte)(102)))));
            this.TopPanel.Controls.Add(this.label8);
            this.TopPanel.Controls.Add(this.FormControl_Panel);
            this.TopPanel.Controls.Add(this.ApplicationName_Label);
            this.TopPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.TopPanel.ForeColor = System.Drawing.Color.Coral;
            this.TopPanel.Location = new System.Drawing.Point(0, 0);
            this.TopPanel.Name = "TopPanel";
            this.TopPanel.Size = new System.Drawing.Size(544, 40);
            this.TopPanel.TabIndex = 0;
            this.TopPanel.MouseDown += new System.Windows.Forms.MouseEventHandler(this.TopPanel_MouseDown);
            this.TopPanel.MouseMove += new System.Windows.Forms.MouseEventHandler(this.TopPanel_MouseMove);
            this.TopPanel.MouseUp += new System.Windows.Forms.MouseEventHandler(this.TopPanel_MouseUp);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.White;
            this.label8.Location = new System.Drawing.Point(204, 14);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(33, 17);
            this.label8.TabIndex = 18;
            this.label8.Text = "v1.2";
            // 
            // FormControl_Panel
            // 
            this.FormControl_Panel.Controls.Add(this.MinimizeLabel);
            this.FormControl_Panel.Controls.Add(this.CloseLabel);
            this.FormControl_Panel.Location = new System.Drawing.Point(483, 5);
            this.FormControl_Panel.Name = "FormControl_Panel";
            this.FormControl_Panel.Size = new System.Drawing.Size(60, 32);
            this.FormControl_Panel.TabIndex = 17;
            // 
            // MinimizeLabel
            // 
            this.MinimizeLabel.AutoSize = true;
            this.MinimizeLabel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.MinimizeLabel.Dock = System.Windows.Forms.DockStyle.Left;
            this.MinimizeLabel.Font = new System.Drawing.Font("Segoe UI Variable Text Semibold", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MinimizeLabel.ForeColor = System.Drawing.Color.White;
            this.MinimizeLabel.Location = new System.Drawing.Point(0, 0);
            this.MinimizeLabel.Name = "MinimizeLabel";
            this.MinimizeLabel.Size = new System.Drawing.Size(31, 26);
            this.MinimizeLabel.TabIndex = 1;
            this.MinimizeLabel.Text = "―";
            this.MinimizeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.MinimizeLabel.Click += new System.EventHandler(this.MinimizeLabel_Click);
            // 
            // CloseLabel
            // 
            this.CloseLabel.AutoSize = true;
            this.CloseLabel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.CloseLabel.Dock = System.Windows.Forms.DockStyle.Right;
            this.CloseLabel.Font = new System.Drawing.Font("Segoe UI Variable Text Semibold", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CloseLabel.ForeColor = System.Drawing.Color.White;
            this.CloseLabel.Location = new System.Drawing.Point(36, 0);
            this.CloseLabel.Name = "CloseLabel";
            this.CloseLabel.Size = new System.Drawing.Size(24, 26);
            this.CloseLabel.TabIndex = 0;
            this.CloseLabel.Text = "X";
            this.CloseLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.CloseLabel.Click += new System.EventHandler(this.CloseLabel_Click);
            // 
            // ApplicationName_Label
            // 
            this.ApplicationName_Label.AutoSize = true;
            this.ApplicationName_Label.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ApplicationName_Label.ForeColor = System.Drawing.Color.White;
            this.ApplicationName_Label.Location = new System.Drawing.Point(9, 6);
            this.ApplicationName_Label.Name = "ApplicationName_Label";
            this.ApplicationName_Label.Size = new System.Drawing.Size(193, 25);
            this.ApplicationName_Label.TabIndex = 0;
            this.ApplicationName_Label.Text = "UPDL Speed Tracker";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(544, 250);
            this.Controls.Add(this.MainPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.MainPanel.ResumeLayout(false);
            this.BottomPanel.ResumeLayout(false);
            this.BottomPanel.PerformLayout();
            this.DOWNLOAD_GroupBox.ResumeLayout(false);
            this.DOWNLOAD_GroupBox.PerformLayout();
            this.UPLOAD_GroupBox.ResumeLayout(false);
            this.UPLOAD_GroupBox.PerformLayout();
            this.TopPanel.ResumeLayout(false);
            this.TopPanel.PerformLayout();
            this.FormControl_Panel.ResumeLayout(false);
            this.FormControl_Panel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel MainPanel;
        private System.Windows.Forms.Panel BottomPanel;
        private System.Windows.Forms.Panel TopPanel;
        private System.Windows.Forms.Label CloseLabel;
        private System.Windows.Forms.Label ApplicationName_Label;
        private System.Windows.Forms.Label MinimizeLabel;
        private System.Windows.Forms.Button SubmitButton;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox UPLOAD_GroupBox;
        private System.Windows.Forms.Panel FormControl_Panel;
        private System.Windows.Forms.GroupBox DOWNLOAD_GroupBox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        internal System.Windows.Forms.DateTimePicker SubmitDate_DatePicker;
        internal System.Windows.Forms.DateTimePicker UPLOADEnd_TimePicker;
        internal System.Windows.Forms.DateTimePicker UPLOADStart_TimePicker;
        internal System.Windows.Forms.DateTimePicker DOWNLOADStart_TimePicker;
        internal System.Windows.Forms.DateTimePicker DOWNLOADEnd_TimePicker;
        internal System.Windows.Forms.ComboBox Cycle_ComboBox;
        internal System.Windows.Forms.TextBox FileSize_TextBox;
        private System.Windows.Forms.Label label8;
    }
}

