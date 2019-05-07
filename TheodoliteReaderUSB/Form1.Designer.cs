namespace TheodoliteReaderUSB
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
            this.buttonSerialPorts = new System.Windows.Forms.Button();
            this.listBoxSerial = new System.Windows.Forms.ListBox();
            this.buttonConnect = new System.Windows.Forms.Button();
            this.textBoxErrorLog = new System.Windows.Forms.TextBox();
            this.labelErrorText = new System.Windows.Forms.Label();
            this.buttonExcelFiles = new System.Windows.Forms.Button();
            this.listBoxExcel = new System.Windows.Forms.ListBox();
            this.buttonLockExcel = new System.Windows.Forms.Button();
            this.radioButtonManual = new System.Windows.Forms.RadioButton();
            this.radioButtonAuto = new System.Windows.Forms.RadioButton();
            this.buttonStartDataCollection = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.buttonGenerate = new System.Windows.Forms.Button();
            this.radioButtonCPEP = new System.Windows.Forms.RadioButton();
            this.radioButtonPEP = new System.Windows.Forms.RadioButton();
            this.radioButtonDEP = new System.Windows.Forms.RadioButton();
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxDirectory = new System.Windows.Forms.TextBox();
            this.textboxVerticalDownFOV = new System.Windows.Forms.TextBox();
            this.textBoxVerticalUpFOV = new System.Windows.Forms.TextBox();
            this.textBoxHorizontalFOV = new System.Windows.Forms.TextBox();
            this.groupBoxManualSettings = new System.Windows.Forms.GroupBox();
            this.checkBoxDist = new System.Windows.Forms.CheckBox();
            this.checkBoxEL = new System.Windows.Forms.CheckBox();
            this.checkBoxAZ = new System.Windows.Forms.CheckBox();
            this.checkBoxZ = new System.Windows.Forms.CheckBox();
            this.checkBoxY = new System.Windows.Forms.CheckBox();
            this.checkBoxX = new System.Windows.Forms.CheckBox();
            this.buttonHelp = new System.Windows.Forms.Button();
            this.radioButtonHorizontalManual = new System.Windows.Forms.RadioButton();
            this.radioButtonVerticalManual = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBoxManualSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonSerialPorts
            // 
            this.buttonSerialPorts.Location = new System.Drawing.Point(5, 104);
            this.buttonSerialPorts.Name = "buttonSerialPorts";
            this.buttonSerialPorts.Size = new System.Drawing.Size(62, 57);
            this.buttonSerialPorts.TabIndex = 0;
            this.buttonSerialPorts.Text = "Refresh Serial Port List";
            this.buttonSerialPorts.UseVisualStyleBackColor = true;
            this.buttonSerialPorts.Click += new System.EventHandler(this.buttonSerialPorts_Click);
            // 
            // listBoxSerial
            // 
            this.listBoxSerial.FormattingEnabled = true;
            this.listBoxSerial.Location = new System.Drawing.Point(5, 19);
            this.listBoxSerial.Name = "listBoxSerial";
            this.listBoxSerial.Size = new System.Drawing.Size(144, 82);
            this.listBoxSerial.TabIndex = 1;
            this.listBoxSerial.SelectedIndexChanged += new System.EventHandler(this.listBoxSerialPorts_SelectedIndexChanged);
            // 
            // buttonConnect
            // 
            this.buttonConnect.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.buttonConnect.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonConnect.Location = new System.Drawing.Point(73, 105);
            this.buttonConnect.Name = "buttonConnect";
            this.buttonConnect.Size = new System.Drawing.Size(57, 55);
            this.buttonConnect.TabIndex = 2;
            this.buttonConnect.Text = "Connect";
            this.buttonConnect.UseVisualStyleBackColor = false;
            this.buttonConnect.Click += new System.EventHandler(this.buttonConnect_Click);
            // 
            // textBoxErrorLog
            // 
            this.textBoxErrorLog.Location = new System.Drawing.Point(12, 316);
            this.textBoxErrorLog.Multiline = true;
            this.textBoxErrorLog.Name = "textBoxErrorLog";
            this.textBoxErrorLog.Size = new System.Drawing.Size(686, 165);
            this.textBoxErrorLog.TabIndex = 3;
            // 
            // labelErrorText
            // 
            this.labelErrorText.AutoSize = true;
            this.labelErrorText.Location = new System.Drawing.Point(9, 300);
            this.labelErrorText.Name = "labelErrorText";
            this.labelErrorText.Size = new System.Drawing.Size(46, 13);
            this.labelErrorText.TabIndex = 4;
            this.labelErrorText.Text = "Info Log";
            // 
            // buttonExcelFiles
            // 
            this.buttonExcelFiles.Location = new System.Drawing.Point(6, 104);
            this.buttonExcelFiles.Name = "buttonExcelFiles";
            this.buttonExcelFiles.Size = new System.Drawing.Size(56, 56);
            this.buttonExcelFiles.TabIndex = 6;
            this.buttonExcelFiles.Text = "Refresh Excel Files";
            this.buttonExcelFiles.UseVisualStyleBackColor = true;
            this.buttonExcelFiles.Click += new System.EventHandler(this.buttonExcelFiles_Click);
            // 
            // listBoxExcel
            // 
            this.listBoxExcel.FormattingEnabled = true;
            this.listBoxExcel.Location = new System.Drawing.Point(3, 17);
            this.listBoxExcel.Name = "listBoxExcel";
            this.listBoxExcel.Size = new System.Drawing.Size(224, 82);
            this.listBoxExcel.TabIndex = 8;
            this.listBoxExcel.SelectedIndexChanged += new System.EventHandler(this.listBoxExcel_SelectedIndexChanged);
            // 
            // buttonLockExcel
            // 
            this.buttonLockExcel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.buttonLockExcel.Location = new System.Drawing.Point(68, 105);
            this.buttonLockExcel.Name = "buttonLockExcel";
            this.buttonLockExcel.Size = new System.Drawing.Size(56, 56);
            this.buttonLockExcel.TabIndex = 9;
            this.buttonLockExcel.Text = "Connect";
            this.buttonLockExcel.UseVisualStyleBackColor = false;
            this.buttonLockExcel.Click += new System.EventHandler(this.buttonLockExcel_Click);
            // 
            // radioButtonManual
            // 
            this.radioButtonManual.AutoSize = true;
            this.radioButtonManual.Location = new System.Drawing.Point(18, 38);
            this.radioButtonManual.Name = "radioButtonManual";
            this.radioButtonManual.Size = new System.Drawing.Size(60, 17);
            this.radioButtonManual.TabIndex = 10;
            this.radioButtonManual.TabStop = true;
            this.radioButtonManual.Text = "Manual";
            this.radioButtonManual.UseVisualStyleBackColor = true;
            this.radioButtonManual.CheckedChanged += new System.EventHandler(this.radioButtonManual_CheckedChanged);
            // 
            // radioButtonAuto
            // 
            this.radioButtonAuto.AutoSize = true;
            this.radioButtonAuto.Location = new System.Drawing.Point(84, 38);
            this.radioButtonAuto.Name = "radioButtonAuto";
            this.radioButtonAuto.Size = new System.Drawing.Size(47, 17);
            this.radioButtonAuto.TabIndex = 11;
            this.radioButtonAuto.TabStop = true;
            this.radioButtonAuto.Text = "Auto";
            this.radioButtonAuto.UseVisualStyleBackColor = true;
            this.radioButtonAuto.CheckedChanged += new System.EventHandler(this.radioButtonAuto_CheckedChanged);
            // 
            // buttonStartDataCollection
            // 
            this.buttonStartDataCollection.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.buttonStartDataCollection.Location = new System.Drawing.Point(6, 39);
            this.buttonStartDataCollection.Name = "buttonStartDataCollection";
            this.buttonStartDataCollection.Size = new System.Drawing.Size(108, 86);
            this.buttonStartDataCollection.TabIndex = 12;
            this.buttonStartDataCollection.Text = "START WORKING";
            this.buttonStartDataCollection.UseVisualStyleBackColor = false;
            this.buttonStartDataCollection.Click += new System.EventHandler(this.buttonStartDataCollection_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButtonAuto);
            this.groupBox1.Controls.Add(this.radioButtonManual);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(7, 173);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(162, 60);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = " STEP 3 - Set Data Collection Mode";
            // 
            // groupBox2
            // 
            this.groupBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.groupBox2.Controls.Add(this.buttonConnect);
            this.groupBox2.Controls.Add(this.listBoxSerial);
            this.groupBox2.Controls.Add(this.buttonSerialPorts);
            this.groupBox2.Location = new System.Drawing.Point(7, 4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(157, 163);
            this.groupBox2.TabIndex = 14;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "STEP 1 - Select Serial Port";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.buttonLockExcel);
            this.groupBox3.Controls.Add(this.listBoxExcel);
            this.groupBox3.Controls.Add(this.buttonExcelFiles);
            this.groupBox3.Location = new System.Drawing.Point(192, 4);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(233, 163);
            this.groupBox3.TabIndex = 15;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "STEP 2 - Selecte Excel File";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.buttonStartDataCollection);
            this.groupBox4.Location = new System.Drawing.Point(300, 173);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(125, 137);
            this.groupBox4.TabIndex = 16;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "STEP 4 - Start Collecting!";
            // 
            // groupBox5
            // 
            this.groupBox5.BackColor = System.Drawing.Color.Wheat;
            this.groupBox5.Controls.Add(this.buttonGenerate);
            this.groupBox5.Controls.Add(this.radioButtonCPEP);
            this.groupBox5.Controls.Add(this.radioButtonPEP);
            this.groupBox5.Controls.Add(this.radioButtonDEP);
            this.groupBox5.Controls.Add(this.buttonBrowse);
            this.groupBox5.Controls.Add(this.label4);
            this.groupBox5.Controls.Add(this.label3);
            this.groupBox5.Controls.Add(this.label2);
            this.groupBox5.Controls.Add(this.label1);
            this.groupBox5.Controls.Add(this.textBoxDirectory);
            this.groupBox5.Controls.Add(this.textboxVerticalDownFOV);
            this.groupBox5.Controls.Add(this.textBoxVerticalUpFOV);
            this.groupBox5.Controls.Add(this.textBoxHorizontalFOV);
            this.groupBox5.Location = new System.Drawing.Point(431, 4);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(267, 280);
            this.groupBox5.TabIndex = 17;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Create Excel Grid (Used for AUTO mode)";
            // 
            // buttonGenerate
            // 
            this.buttonGenerate.Location = new System.Drawing.Point(80, 230);
            this.buttonGenerate.Name = "buttonGenerate";
            this.buttonGenerate.Size = new System.Drawing.Size(117, 43);
            this.buttonGenerate.TabIndex = 12;
            this.buttonGenerate.Text = "Create!";
            this.buttonGenerate.UseVisualStyleBackColor = true;
            this.buttonGenerate.Click += new System.EventHandler(this.buttonGenerate_Click);
            // 
            // radioButtonCPEP
            // 
            this.radioButtonCPEP.AutoSize = true;
            this.radioButtonCPEP.Location = new System.Drawing.Point(125, 208);
            this.radioButtonCPEP.Name = "radioButtonCPEP";
            this.radioButtonCPEP.Size = new System.Drawing.Size(53, 17);
            this.radioButtonCPEP.TabIndex = 11;
            this.radioButtonCPEP.TabStop = true;
            this.radioButtonCPEP.Text = "CPEP";
            this.radioButtonCPEP.UseVisualStyleBackColor = true;
            this.radioButtonCPEP.CheckedChanged += new System.EventHandler(this.radioButtonCPEP_CheckedChanged);
            // 
            // radioButtonPEP
            // 
            this.radioButtonPEP.AutoSize = true;
            this.radioButtonPEP.Location = new System.Drawing.Point(73, 208);
            this.radioButtonPEP.Name = "radioButtonPEP";
            this.radioButtonPEP.Size = new System.Drawing.Size(46, 17);
            this.radioButtonPEP.TabIndex = 10;
            this.radioButtonPEP.TabStop = true;
            this.radioButtonPEP.Text = "PEP";
            this.radioButtonPEP.UseVisualStyleBackColor = true;
            this.radioButtonPEP.CheckedChanged += new System.EventHandler(this.radioButtonPEP_CheckedChanged);
            // 
            // radioButtonDEP
            // 
            this.radioButtonDEP.AutoSize = true;
            this.radioButtonDEP.Location = new System.Drawing.Point(20, 208);
            this.radioButtonDEP.Name = "radioButtonDEP";
            this.radioButtonDEP.Size = new System.Drawing.Size(47, 17);
            this.radioButtonDEP.TabIndex = 9;
            this.radioButtonDEP.TabStop = true;
            this.radioButtonDEP.Text = "DEP";
            this.radioButtonDEP.UseVisualStyleBackColor = true;
            this.radioButtonDEP.CheckedChanged += new System.EventHandler(this.radioButtonDEP_CheckedChanged);
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Location = new System.Drawing.Point(19, 163);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(73, 27);
            this.buttonBrowse.TabIndex = 8;
            this.buttonBrowse.Text = "Browse";
            this.buttonBrowse.UseVisualStyleBackColor = true;
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 121);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(76, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Save Location";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(76, 96);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(97, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Vertical FOV Down";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(76, 67);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Vertical FOV UP";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(77, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Horizontal FOV";
            // 
            // textBoxDirectory
            // 
            this.textBoxDirectory.Enabled = false;
            this.textBoxDirectory.Location = new System.Drawing.Point(19, 137);
            this.textBoxDirectory.Name = "textBoxDirectory";
            this.textBoxDirectory.ReadOnly = true;
            this.textBoxDirectory.Size = new System.Drawing.Size(242, 20);
            this.textBoxDirectory.TabIndex = 3;
            // 
            // textboxVerticalDownFOV
            // 
            this.textboxVerticalDownFOV.Location = new System.Drawing.Point(20, 93);
            this.textboxVerticalDownFOV.Name = "textboxVerticalDownFOV";
            this.textboxVerticalDownFOV.Size = new System.Drawing.Size(51, 20);
            this.textboxVerticalDownFOV.TabIndex = 2;
            // 
            // textBoxVerticalUpFOV
            // 
            this.textBoxVerticalUpFOV.Location = new System.Drawing.Point(20, 64);
            this.textBoxVerticalUpFOV.Name = "textBoxVerticalUpFOV";
            this.textBoxVerticalUpFOV.Size = new System.Drawing.Size(50, 20);
            this.textBoxVerticalUpFOV.TabIndex = 1;
            // 
            // textBoxHorizontalFOV
            // 
            this.textBoxHorizontalFOV.Location = new System.Drawing.Point(19, 31);
            this.textBoxHorizontalFOV.Name = "textBoxHorizontalFOV";
            this.textBoxHorizontalFOV.Size = new System.Drawing.Size(52, 20);
            this.textBoxHorizontalFOV.TabIndex = 0;
            // 
            // groupBoxManualSettings
            // 
            this.groupBoxManualSettings.Controls.Add(this.radioButtonVerticalManual);
            this.groupBoxManualSettings.Controls.Add(this.radioButtonHorizontalManual);
            this.groupBoxManualSettings.Controls.Add(this.checkBoxDist);
            this.groupBoxManualSettings.Controls.Add(this.checkBoxEL);
            this.groupBoxManualSettings.Controls.Add(this.checkBoxAZ);
            this.groupBoxManualSettings.Controls.Add(this.checkBoxZ);
            this.groupBoxManualSettings.Controls.Add(this.checkBoxY);
            this.groupBoxManualSettings.Controls.Add(this.checkBoxX);
            this.groupBoxManualSettings.Location = new System.Drawing.Point(7, 239);
            this.groupBoxManualSettings.Name = "groupBoxManualSettings";
            this.groupBoxManualSettings.Size = new System.Drawing.Size(287, 59);
            this.groupBoxManualSettings.TabIndex = 18;
            this.groupBoxManualSettings.TabStop = false;
            this.groupBoxManualSettings.Text = "Manual Mode Data";
            // 
            // checkBoxDist
            // 
            this.checkBoxDist.AutoSize = true;
            this.checkBoxDist.Location = new System.Drawing.Point(176, 25);
            this.checkBoxDist.Name = "checkBoxDist";
            this.checkBoxDist.Size = new System.Drawing.Size(44, 17);
            this.checkBoxDist.TabIndex = 5;
            this.checkBoxDist.Text = "Dist";
            this.checkBoxDist.UseVisualStyleBackColor = true;
            // 
            // checkBoxEL
            // 
            this.checkBoxEL.AutoSize = true;
            this.checkBoxEL.Location = new System.Drawing.Point(140, 25);
            this.checkBoxEL.Name = "checkBoxEL";
            this.checkBoxEL.Size = new System.Drawing.Size(39, 17);
            this.checkBoxEL.TabIndex = 4;
            this.checkBoxEL.Text = "EL";
            this.checkBoxEL.UseVisualStyleBackColor = true;
            // 
            // checkBoxAZ
            // 
            this.checkBoxAZ.AutoSize = true;
            this.checkBoxAZ.Location = new System.Drawing.Point(103, 25);
            this.checkBoxAZ.Name = "checkBoxAZ";
            this.checkBoxAZ.Size = new System.Drawing.Size(40, 17);
            this.checkBoxAZ.TabIndex = 3;
            this.checkBoxAZ.Text = "AZ";
            this.checkBoxAZ.UseVisualStyleBackColor = true;
            // 
            // checkBoxZ
            // 
            this.checkBoxZ.AutoSize = true;
            this.checkBoxZ.Location = new System.Drawing.Point(72, 25);
            this.checkBoxZ.Name = "checkBoxZ";
            this.checkBoxZ.Size = new System.Drawing.Size(33, 17);
            this.checkBoxZ.TabIndex = 2;
            this.checkBoxZ.Text = "Z";
            this.checkBoxZ.UseVisualStyleBackColor = true;
            // 
            // checkBoxY
            // 
            this.checkBoxY.AutoSize = true;
            this.checkBoxY.Location = new System.Drawing.Point(41, 25);
            this.checkBoxY.Name = "checkBoxY";
            this.checkBoxY.Size = new System.Drawing.Size(33, 17);
            this.checkBoxY.TabIndex = 1;
            this.checkBoxY.Text = "Y";
            this.checkBoxY.UseVisualStyleBackColor = true;
            // 
            // checkBoxX
            // 
            this.checkBoxX.AutoSize = true;
            this.checkBoxX.Location = new System.Drawing.Point(10, 25);
            this.checkBoxX.Name = "checkBoxX";
            this.checkBoxX.Size = new System.Drawing.Size(33, 17);
            this.checkBoxX.TabIndex = 0;
            this.checkBoxX.Text = "X";
            this.checkBoxX.UseVisualStyleBackColor = true;
            // 
            // buttonHelp
            // 
            this.buttonHelp.BackColor = System.Drawing.Color.Transparent;
            this.buttonHelp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonHelp.Location = new System.Drawing.Point(529, 290);
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Size = new System.Drawing.Size(75, 23);
            this.buttonHelp.TabIndex = 19;
            this.buttonHelp.Text = "HELP";
            this.buttonHelp.UseVisualStyleBackColor = false;
            this.buttonHelp.Click += new System.EventHandler(this.buttonHelp_Click);
            // 
            // radioButtonHorizontalManual
            // 
            this.radioButtonHorizontalManual.AutoSize = true;
            this.radioButtonHorizontalManual.Location = new System.Drawing.Point(219, 10);
            this.radioButtonHorizontalManual.Name = "radioButtonHorizontalManual";
            this.radioButtonHorizontalManual.Size = new System.Drawing.Size(72, 17);
            this.radioButtonHorizontalManual.TabIndex = 6;
            this.radioButtonHorizontalManual.TabStop = true;
            this.radioButtonHorizontalManual.Text = "Horizontal";
            this.radioButtonHorizontalManual.UseVisualStyleBackColor = true;
            this.radioButtonHorizontalManual.CheckedChanged += new System.EventHandler(this.radioButtonHorizontalManual_CheckedChanged);
            // 
            // radioButtonVerticalManual
            // 
            this.radioButtonVerticalManual.AutoSize = true;
            this.radioButtonVerticalManual.Location = new System.Drawing.Point(219, 39);
            this.radioButtonVerticalManual.Name = "radioButtonVerticalManual";
            this.radioButtonVerticalManual.Size = new System.Drawing.Size(60, 17);
            this.radioButtonVerticalManual.TabIndex = 7;
            this.radioButtonVerticalManual.TabStop = true;
            this.radioButtonVerticalManual.Text = "Vertical";
            this.radioButtonVerticalManual.UseVisualStyleBackColor = true;
            this.radioButtonVerticalManual.CheckedChanged += new System.EventHandler(this.radioButtonVerticalManual_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(719, 493);
            this.Controls.Add(this.buttonHelp);
            this.Controls.Add(this.groupBoxManualSettings);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.labelErrorText);
            this.Controls.Add(this.textBoxErrorLog);
            this.Name = "Form1";
            this.Text = "TS02 USB Data Collector";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBoxManualSettings.ResumeLayout(false);
            this.groupBoxManualSettings.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonSerialPorts;
        private System.Windows.Forms.ListBox listBoxSerial;
        private System.Windows.Forms.Button buttonConnect;
        private System.Windows.Forms.TextBox textBoxErrorLog;
        private System.Windows.Forms.Label labelErrorText;
        private System.Windows.Forms.Button buttonExcelFiles;
        private System.Windows.Forms.ListBox listBoxExcel;
        private System.Windows.Forms.Button buttonLockExcel;
        private System.Windows.Forms.RadioButton radioButtonManual;
        private System.Windows.Forms.RadioButton radioButtonAuto;
        private System.Windows.Forms.Button buttonStartDataCollection;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxDirectory;
        private System.Windows.Forms.TextBox textboxVerticalDownFOV;
        private System.Windows.Forms.TextBox textBoxVerticalUpFOV;
        private System.Windows.Forms.TextBox textBoxHorizontalFOV;
        private System.Windows.Forms.RadioButton radioButtonCPEP;
        private System.Windows.Forms.RadioButton radioButtonPEP;
        private System.Windows.Forms.RadioButton radioButtonDEP;
        private System.Windows.Forms.Button buttonGenerate;
        private System.Windows.Forms.GroupBox groupBoxManualSettings;
        private System.Windows.Forms.CheckBox checkBoxDist;
        private System.Windows.Forms.CheckBox checkBoxEL;
        private System.Windows.Forms.CheckBox checkBoxAZ;
        private System.Windows.Forms.CheckBox checkBoxZ;
        private System.Windows.Forms.CheckBox checkBoxY;
        private System.Windows.Forms.CheckBox checkBoxX;
        private System.Windows.Forms.Button buttonHelp;
        private System.Windows.Forms.RadioButton radioButtonVerticalManual;
        private System.Windows.Forms.RadioButton radioButtonHorizontalManual;
    }
}

