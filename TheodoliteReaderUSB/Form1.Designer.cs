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
            this.SuspendLayout();
            // 
            // buttonSerialPorts
            // 
            this.buttonSerialPorts.Location = new System.Drawing.Point(15, 175);
            this.buttonSerialPorts.Name = "buttonSerialPorts";
            this.buttonSerialPorts.Size = new System.Drawing.Size(56, 56);
            this.buttonSerialPorts.TabIndex = 0;
            this.buttonSerialPorts.Text = "Refresh Serial Port List";
            this.buttonSerialPorts.UseVisualStyleBackColor = true;
            this.buttonSerialPorts.Click += new System.EventHandler(this.buttonSerialPorts_Click);
            // 
            // listBoxSerial
            // 
            this.listBoxSerial.FormattingEnabled = true;
            this.listBoxSerial.Location = new System.Drawing.Point(15, 74);
            this.listBoxSerial.Name = "listBoxSerial";
            this.listBoxSerial.Size = new System.Drawing.Size(120, 95);
            this.listBoxSerial.TabIndex = 1;
            this.listBoxSerial.SelectedIndexChanged += new System.EventHandler(this.listBoxSerialPorts_SelectedIndexChanged);
            // 
            // buttonConnect
            // 
            this.buttonConnect.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.buttonConnect.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonConnect.Location = new System.Drawing.Point(78, 176);
            this.buttonConnect.Name = "buttonConnect";
            this.buttonConnect.Size = new System.Drawing.Size(57, 55);
            this.buttonConnect.TabIndex = 2;
            this.buttonConnect.Text = "Connect";
            this.buttonConnect.UseVisualStyleBackColor = false;
            this.buttonConnect.Click += new System.EventHandler(this.buttonConnect_Click);
            // 
            // textBoxErrorLog
            // 
            this.textBoxErrorLog.Location = new System.Drawing.Point(32, 349);
            this.textBoxErrorLog.Multiline = true;
            this.textBoxErrorLog.Name = "textBoxErrorLog";
            this.textBoxErrorLog.Size = new System.Drawing.Size(736, 89);
            this.textBoxErrorLog.TabIndex = 3;
            // 
            // labelErrorText
            // 
            this.labelErrorText.AutoSize = true;
            this.labelErrorText.Location = new System.Drawing.Point(29, 333);
            this.labelErrorText.Name = "labelErrorText";
            this.labelErrorText.Size = new System.Drawing.Size(46, 13);
            this.labelErrorText.TabIndex = 4;
            this.labelErrorText.Text = "Info Log";
            // 
            // buttonExcelFiles
            // 
            this.buttonExcelFiles.Location = new System.Drawing.Point(221, 177);
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
            this.listBoxExcel.Location = new System.Drawing.Point(220, 74);
            this.listBoxExcel.Name = "listBoxExcel";
            this.listBoxExcel.Size = new System.Drawing.Size(120, 95);
            this.listBoxExcel.TabIndex = 8;
            this.listBoxExcel.SelectedIndexChanged += new System.EventHandler(this.listBoxExcel_SelectedIndexChanged);
            // 
            // buttonLockExcel
            // 
            this.buttonLockExcel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.buttonLockExcel.Location = new System.Drawing.Point(284, 177);
            this.buttonLockExcel.Name = "buttonLockExcel";
            this.buttonLockExcel.Size = new System.Drawing.Size(56, 56);
            this.buttonLockExcel.TabIndex = 9;
            this.buttonLockExcel.Text = "Connect";
            this.buttonLockExcel.UseVisualStyleBackColor = false;
            this.buttonLockExcel.Click += new System.EventHandler(this.buttonLockExcel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.buttonLockExcel);
            this.Controls.Add(this.listBoxExcel);
            this.Controls.Add(this.buttonExcelFiles);
            this.Controls.Add(this.labelErrorText);
            this.Controls.Add(this.textBoxErrorLog);
            this.Controls.Add(this.buttonConnect);
            this.Controls.Add(this.listBoxSerial);
            this.Controls.Add(this.buttonSerialPorts);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
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
    }
}

