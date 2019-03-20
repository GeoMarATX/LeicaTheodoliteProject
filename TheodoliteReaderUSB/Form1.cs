using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using Excel = Microsoft.Office.Interop.Excel;


namespace TheodoliteReaderUSB
{
    public partial class Form1 : Form
    {
        private BindingList<string> ports;
        private SerialPort serial;
        private string[] portsArray;
        private string portName;

        private const int _baudRate = 1152200;
        private const int _dataBits = 8;
        private const int _readTimeOut = 500;
        private const int _writeTimeOut = 500;
        private const StopBits _stopBits = StopBits.One;
        private const Handshake _handshake = Handshake.None;

        private Excel.Application xlApp;                        //Remember to CLOSE EXCEL!
        private List<Excel.Workbook> workbooks;
        private static Excel.Workbook xlWorkBook;
        private BindingList<String> workbooksStrings;

        static Color myGreen = Color.FromArgb(0, 193, 0);
        static Color myRed = Color.FromArgb(192, 0, 0);

        public Form1()
        {
            InitializeComponent();
        }

        private void buttonSerialPorts_Click(object sender, EventArgs e)
        {
            //Refresh the list of names of connected serial ports.
            ports.Clear();
            portsArray = SerialPort.GetPortNames();

            //Add each port name from a string array to a list of strings.
            foreach (string myString in portsArray)
            {
                ports.Add(myString);
            }
            if(ports.Count == 0)
            {
                textBoxErrorLog.AppendText(string.Format("No serial devices found{0}", Environment.NewLine));
            }
        }

        private void buttonConnect_Click(object sender, EventArgs e)
        {
            if(buttonConnect.Text == "Connect")
            {
                //When button text is "Connect", perform the following actions to connect the serial port.
                if (!string.IsNullOrWhiteSpace(portName))
                {
                    try
                    {
                        //Init serial port parameters.
                        initSerialPort();

                        //The 'serial' object instance is created in Form1_Load().
                        serial.Open();

                        if (serial.IsOpen)
                        {
                            //Serial port has been opened succesfully. Now change the "Connect" button to display "Disconnect".
                            buttonConnect.Font = new Font(buttonConnect.Font.FontFamily, 6);        //Font size defaults to 8, which is too large to display nicely.
                            buttonConnect.Text = "Disconnect";
                            buttonConnect.BackColor = myRed;

                            //Hide the button which refreshes the list of serial ports.
                            buttonSerialPorts.Hide();

                            //Disable the window which shows the list of currenlty open Excel files.
                            listBoxSerial.Enabled = false;

                            //Display message to the log window indicating that the serial port has been opened.
                            //textBoxErrorLog.AppendText(string.Format("Connected to {0}.{1}", portName, Environment.NewLine));
                            myLog(string.Format("Connected to {0}", portName));
                        }
                    }
                    catch (Exception ex) when (ex is UnauthorizedAccessException || ex is ArgumentOutOfRangeException || ex is ArgumentException || ex is InvalidOperationException || ex is ArgumentNullException || ex is System.IO.IOException)
                    {
                        //Write execption descriptions to the log window.
                        //textBoxErrorLog.AppendText(string.Format("buttonConnect_Click() :: {0}", ex.Message));
                        myLog(string.Format("buttonConnect_Click() :: {0}", ex.Message));

                    }
                }
                else
                {
                    //If portName was never set (blank, null or white space), display an error message.
                    //textBoxErrorLog.AppendText(string.Format("Connect failed, no serial port selected{0}", Environment.NewLine));
                    myLog(string.Format("Connect failed, no serial port selected"));
                }
            }
            else if(buttonConnect.Text == "Disconnect")
            {
                //When buttonConnect is displaying the text "Disconnect", the serial object is connected and open.
                try
                {
                    //Perfrom tasks related to closing the serial port.
                    serial.Close();

                    if(!serial.IsOpen)
                    {
                        //Serial closed successfully. Reset portName.
                        portName = "";

                        //Change the apperance of the 'Connect' button.
                        buttonConnect.Font = new Font(buttonConnect.Font.FontFamily, 8);
                        buttonConnect.Text = "Connect";
                        buttonConnect.BackColor = myGreen;

                        //Un hide the button which refreshes the list of serial ports.
                        buttonSerialPorts.Show();

                        //Re-enable the window which shows the list of currenlty open Excel files.
                        listBoxSerial.Enabled = true;

                        //Display message to the log window indicating that the serial port has been closed.
                        //textBoxErrorLog.AppendText(string.Format("Serial closed {0}", Environment.NewLine));
                        myLog(string.Format("Serial closed "));

                    }
                }
                catch (Exception ex) when (ex is System.IO.IOException || ex is ArgumentNullException || ex is ArgumentException)
                {
                    //Write execption descriptions to the log window.
                    //textBoxErrorLog.AppendText(string.Format("buttonConnect_Click() :: {0}{1}", ex.Message, Environment.NewLine));
                    myLog(string.Format("buttonConnect_Click() :: {0}", ex.Message));

                }
            }
            
            
        }

        private void buttonExcelFiles_Click(object sender, EventArgs e)
        {
            //Reset List<Excel.Workbook> workbooks
            workbooks.Clear();
            workbooksStrings.Clear();

            xlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

            foreach (Excel.Workbook item in xlApp.Workbooks)
            {
                workbooks.Add(item);
                workbooksStrings.Add(item.Name);
            }

            if (workbooks.Count == 0)
            {
                myLog(string.Format("No Excel files found"));
            }
        }

        private void buttonLockExcel_Click(object sender, EventArgs e)
        {

            if (buttonLockExcel.Text == "Connect")
            {
                //The user has requested to connect to a selected Excel workbook.
                if (xlWorkBook != null)
                {
                    //Change the "Connect" button text to "Disconnect".
                    buttonLockExcel.Font = new Font(buttonLockExcel.Font.FontFamily, 6);        //Font size defaults to 8, which is too large to display nicely.
                    buttonLockExcel.Text = "Disconnect";
                    buttonLockExcel.BackColor = myRed;

                    //Disable the listbox which lists the currently open Excel files.
                    listBoxExcel.Enabled = false;

                    //Hide the "Refresh Excel Files" button.
                    buttonExcelFiles.Visible = false;

                    if (xlWorkBook != null)
                    {
                        myLog(string.Format("Connected to Excel sheet {0}", xlWorkBook.Name));
                    }
                }
                else
                {
                    myLog(string.Format("Error: Cannot connect to Excel workbook. No workbook selected."));
                }
            }
            else if (buttonLockExcel.Text == "Disconnect")
            {
                //Change the "Disconnect" button text to "Connect".
                buttonLockExcel.Font = new Font(buttonLockExcel.Font.FontFamily, 8);        //Font size defaults to 8, which is too large to display nicely.
                buttonLockExcel.Text = "Connect";
                buttonLockExcel.BackColor = myGreen;


                //Enable the listbox which lists the currently open Excel files.
                listBoxExcel.Enabled = true;

                //Show the "Refresh Excel Files" button.
                buttonExcelFiles.Visible = true;

                try
                {
                    //This is to check that the user has not closed the Excel file before disconnecting.
                    //Future work: Trigger event to handle user closing file before disconnecting.
                    if (xlWorkBook != null)
                    {
                        myLog(string.Format("Disconnected from workbook {0}", xlWorkBook.Name));
                    }
                    else
                    {
                        myLog(string.Format("Disconnected from workbook. Workbook = null"));
                    }
                    xlWorkBook = null;
                }
                catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException)
                {
                    myLog(string.Format("buttonLockEcel_Click :: {0}", ex.Message));
                    xlWorkBook = null;                  //Something happened to the work book. 
                    buttonExcelFiles.PerformClick();    //Refresh the list of excel files. User most likely closed the file before disconnecting.
                }
            }

        }
        
        private void listBoxSerialPorts_SelectedIndexChanged(object sender, EventArgs e)
        {
            portName = (string)listBoxSerial.SelectedValue;

        }

        private void listBoxExcel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxExcel.SelectedIndex < 0)
            {
                listBoxExcel.ClearSelected();
            }
            else
            {
                xlWorkBook = workbooks.ElementAt(listBoxExcel.SelectedIndex);                
            }           
        }
        
        private void initSerialPort()
        {
            //Set serial port parameters
            serial.PortName = portName;
            serial.BaudRate = _baudRate;
            serial.DataBits = _dataBits;
            serial.StopBits = _stopBits;
            serial.Handshake = _handshake;

            // Set the read/write timeouts
            serial.ReadTimeout = _readTimeOut;
            serial.WriteTimeout = _writeTimeOut;

        }
        
        private void myLog(string log)
        {
            textBoxErrorLog.AppendText(log);
            if(!log.EndsWith(Environment.NewLine))
            {
                textBoxErrorLog.AppendText(Environment.NewLine);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            /*
             *  Serial port inits.
             */
            serial = new SerialPort();
            ports = new BindingList<string>();                   //Initialize the list of serialport names, stored as strings.           
            listBoxSerial.DataSource = ports;                    //Bind the list 'ports' to the list box 'listBoxSerialPorts'.

            /*
             * Excel inits.
             */
            xlApp = new Excel.Application();
            workbooks = new List<Excel.Workbook>();             //"Refresh Excel Files" button refreshes the var 'workbooks' with the currently open excel files.
            workbooksStrings = new BindingList<String>();       //Workaround because List<Excel.Workbook> does not display name strings properly when data bound.
            listBoxExcel.DataSource = workbooksStrings;
            buttonLockExcel.BackColor = myGreen;                //Init button color to myGreen, defined at top. 
            xlApp.EnableEvents = true;
            //xlApp.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(XlApp_WorkbookBeforeClose);
            xlApp.Visible = true;
        }

    }
}


/*
        private void XlApp_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            Excel.Workbook myWorkbook = Wb;
            DialogResult result = MessageBox.Show("Are you sure you want to close? " +
            "Theodolite is currently connected to this file. " + this.Name + "?", "Example",
            MessageBoxButtons.YesNoCancel);

            switch (result)
            {
                case DialogResult.Yes:
                    Wb.Save();
                    break;

                case DialogResult.Cancel:
                    Cancel = true;
                    break;

                // The following code ensures that the default Save File 
                // dialog is not displayed.
                case DialogResult.No:
                    Wb.Saved = true;
                    break;
            }
        }
 */
