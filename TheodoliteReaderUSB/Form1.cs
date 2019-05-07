using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel; 


namespace TheodoliteReaderUSB
{
    public partial class Form1 : Form
    {
        private delegate void SafeCallDelegate(string text);

        private BindingList<string> ports;
        private static SerialPort serial;
        private string[] portsArray;
        private string portName;

        private const int _baudRate = 115200;
        private const int _dataBits = 8;
        private const int _readTimeOut = 500;
        private const int _writeTimeOut = 500;
        private const StopBits _stopBits = StopBits.One;
        private const Handshake _handshake = Handshake.None;

        private static Excel.Application xlApp;                        
        private List<Excel.Workbook> workbooks;
        private static Excel.Workbook xlWorkBook;
        private BindingList<String> workbooksStrings;

        static int horizontalFOV;
        static int verticalUpFOV;
        static int verticalDownFOV;
        static string saveFileDir;
        static string saveFile;
        static string eyepoint;

        static string DATA_COLLECTION_SHEET_NAME = "Data Collection";
        static string TOTAL_DISTORTION_SHEET_NAME = "Total Dist";
        static string RELATIVE_DISTORTION_SHEET_NAME = "Rel Dist";

        static Color myGreen = Color.FromArgb(0, 193, 0);
        static Color myRed = Color.FromArgb(192, 0, 0);
        static Color myGrey = Color.FromArgb(191, 191, 191);

        private static bool autoMode;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (string.Equals((sender as Button).Name, @"CloseButton"))
            {
                // Do something proper to CloseButton.
            }
            else
            {
                // Then assume that X has been clicked and act accordingly.
                xlWorkBook.Close(0);
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            /*
             *  Serial port inits.
             */
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

            checkBoxX.Checked = true;
            checkBoxY.Checked = true;
            checkBoxZ.Checked = true;
            checkBoxAZ.Checked = true;
            checkBoxEL.Checked = true;
            checkBoxDist.Checked = true;

            radioButtonManual.Checked = true;
            radioButtonAuto.Checked = false;

            radioButtonHorizontalManual.Checked = true;
            radioButtonVerticalManual.Checked = false;
            
        }
        
        /*
         * Serial Port Related
         */
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
            else
            {
                listBoxSerial.SetSelected(0, true);
            }
        }

        private void buttonConnect_Click(object sender, EventArgs e)
        {
            if(buttonConnect.Text == "Connect")
            {
                //When button text is "Connect", the serial port is not connected. Perform the following actions to connect the serial port.
                if (!string.IsNullOrWhiteSpace(portName))
                {
                    try
                    {
                        serial = new SerialPort();

                        //Initialize the serial port parameters.
                        initSerialPort();

                        //The 'serial' object instance is created in Form1_Load().
                        serial.Open();

                        if (serial.IsOpen)
                        {
                            //Serial port has been opened successfully. Now change the "Connect" button to display "Disconnect".
                            buttonConnect.Font = new Font(buttonConnect.Font.FontFamily, 6);        //Font size defaults to 8, which is too large to display nicely.
                            buttonConnect.Text = "Disconnect";
                            buttonConnect.BackColor = myRed;

                            //Hide the button that is used to refresh the list of serial ports.
                            buttonSerialPorts.Hide();

                            //Disable the window which shows the list of currently open Excel files.
                            listBoxSerial.Enabled = false;

                            //Display message to the log window indicating that the serial port has been opened.
                            //textBoxErrorLog.AppendText(string.Format("Connected to {0}.{1}", portName, Environment.NewLine));
                            myLog(string.Format("Connected to {0}", portName));

                            //Enable event handler.
                            //serial.DataReceived += serialPort_DataReceived;

                        }
                    }
                    catch (Exception ex) when (ex is UnauthorizedAccessException || ex is ArgumentOutOfRangeException || ex is ArgumentException || ex is InvalidOperationException || ex is ArgumentNullException || ex is System.IO.IOException)
                    {
                        //Write exception descriptions to the log window.
                        myLog(string.Format("buttonConnect_Click() :: {0}", ex.Message));

                    }
                }
                else
                {
                    //If portName was never set (blank, null or white space), display an error message.
                    myLog(string.Format("Connect failed, no serial port selected"));
                }
            }
            else if(buttonConnect.Text == "Disconnect")
            {
                //When buttonConnect is displaying the text "Disconnect", the serial object is connected and open. Perform the following step to disconnect the serial port.
                try
                {
                    //Close the serial port.
                    serial.Close();

                    if(!serial.IsOpen)
                    {
                        //Serial port was closed successfully. Reset portName.
                        portName = "";

                        //Change the appearance of the 'Connect' button.
                        buttonConnect.Font = new Font(buttonConnect.Font.FontFamily, 8);
                        buttonConnect.Text = "Connect";
                        buttonConnect.BackColor = myGreen;

                        //Un-hide the button which refreshes the list of serial ports.
                        buttonSerialPorts.Show();

                        //Re-enable the window which shows the list of currently open Excel files.
                        listBoxSerial.Enabled = true;

                        //Display message to the log window indicating that the serial port has been closed.
                        //textBoxErrorLog.AppendText(string.Format("Serial closed {0}", Environment.NewLine));
                        myLog(string.Format("Serial closed "));

                        //Disable event handler.
                        //serial.DataReceived -= serialPort_DataReceived;
                    }
                }
                catch (Exception ex) when (ex is System.IO.IOException || ex is ArgumentNullException || ex is ArgumentException)
                {
                    //Write exception descriptions to the log window.
                    //textBoxErrorLog.AppendText(string.Format("buttonConnect_Click() :: {0}{1}", ex.Message, Environment.NewLine));
                    myLog(string.Format("buttonConnect_Click() :: {0}", ex.Message));

                }
            }
            
            
        }

        private void serialPort_DataReceived_Manual(object sender, SerialDataReceivedEventArgs e)
        {
            var sp = sender as SerialPort;

            var gsiParser = new ParseGSI();

            try
            {
                //Try reading from the serial port.
                gsiParser.Raw = sp.ReadLine();
            }
            catch (Exception ex) when (ex is InvalidOperationException || ex is TimeoutException)
            {
                myLog(string.Format("buttonStartDataCollection_Click :: {0}", ex.Message));
            }

            //Store to active excell sheet. Advance the cursor after storing each value.
            Excel.Range rng = xlApp.ActiveCell;
            Excel.Worksheet sheet = xlWorkBook.ActiveSheet;
            int count = 0;

            if(radioButtonHorizontalManual.Checked == true)
            {
                if (checkBoxX.Checked == true)
                {
                    rng.Cells.Value = gsiParser.TargetEasting;
                    rng = (Excel.Range)sheet.Cells[rng.Row, rng.Column + 1];
                    count++;
                }
                if (checkBoxY.Checked == true)
                {
                    rng.Cells.Value = gsiParser.TargetNorthing;
                    rng = (Excel.Range)sheet.Cells[rng.Row, rng.Column + 1];
                    count++;
                }
                if (checkBoxZ.Checked == true)
                {
                    rng.Cells.Value = gsiParser.TargetElevation;
                    rng = (Excel.Range)sheet.Cells[rng.Row, rng.Column + 1];
                    count++;
                }
                if (checkBoxAZ.Checked == true)
                {
                    rng.Cells.Value = gsiParser.HorizontalAngle;
                    rng = (Excel.Range)sheet.Cells[rng.Row, rng.Column + 1];
                    count++;
                }
                if (checkBoxEL.Checked == true)
                {
                    rng.Cells.Value = gsiParser.VerticalAngle;
                    rng = (Excel.Range)sheet.Cells[rng.Row, rng.Column + 1];
                    count++;
                }
                if (checkBoxDist.Checked == true)
                {
                    rng.Cells.Value = gsiParser.HorizontalDistance;

                }

                rng = (Excel.Range)sheet.Cells[rng.Row + 1, rng.Column - count];
                rng.Select();
            }
            if (radioButtonVerticalManual.Checked == true)
            {
                if (checkBoxX.Checked == true)
                {
                    rng.Cells.Value = gsiParser.TargetEasting;
                    rng = (Excel.Range)sheet.Cells[rng.Row + 1, rng.Column];
                    count++;
                }
                if (checkBoxY.Checked == true)
                {
                    rng.Cells.Value = gsiParser.TargetNorthing;
                    rng = (Excel.Range)sheet.Cells[rng.Row + 1, rng.Column];
                    count++;
                }
                if (checkBoxZ.Checked == true)
                {
                    rng.Cells.Value = gsiParser.TargetElevation;
                    rng = (Excel.Range)sheet.Cells[rng.Row + 1, rng.Column];
                    count++;
                }
                if (checkBoxAZ.Checked == true)
                {
                    rng.Cells.Value = gsiParser.HorizontalAngle;
                    rng = (Excel.Range)sheet.Cells[rng.Row + 1, rng.Column];
                    count++;
                }
                if (checkBoxEL.Checked == true)
                {
                    rng.Cells.Value = gsiParser.VerticalAngle;
                    rng = (Excel.Range)sheet.Cells[rng.Row + 1, rng.Column];
                    count++;
                }
                if (checkBoxDist.Checked == true)
                {
                    rng.Cells.Value = gsiParser.HorizontalDistance;

                }

                rng = (Excel.Range)sheet.Cells[rng.Row, rng.Column];
                rng.Select();
            }

        }

        private void serialPort_DataReceived_Auto(object sender, SerialDataReceivedEventArgs e)
        {
            var sp = sender as SerialPort;
            var gsiParser = new ParseGSI(HorizontalAngleMode.PLUS_MINUS_180,VerticalAngleMode.PLUS_MINUS_180);
            try
            {
                //Try reading from the serial port.
                gsiParser.Raw = sp.ReadLine();
            }
            catch (Exception ex) when (ex is InvalidOperationException || ex is TimeoutException)
            {
                myLog(string.Format("buttonStartDataCollection_Click :: {0}", ex.Message));
            }           
            
            try
            {
                //Data collection area starts at row 3, col 2.
                var rowOffset = 3;
                var colOffset = 2;

                var endFound = false;
                var value = (int)xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[rowOffset, colOffset - 1].Value;
                var myVertUpFOV = value;
                var myVertDownFOV = 0;
                var myHorzLeft = 0;
                var myHorzRight = 0;
                //Check that connect Excel sheet is formated correctly, elevation.
                for (int i = 0; !endFound; i++)
                {
                    if (xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[rowOffset + i * 2 + 2, colOffset - 1].Value2 != null)
                    {
                        if ((xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[rowOffset + i * 2 + 2, colOffset - 1].Value2 == value - 5))
                        {
                            value = (int)xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[rowOffset + i * 2 + 2, colOffset - 1].Value2;
                        }
                    }
                    else
                    {
                        myVertDownFOV = value;
                        endFound = true;
                    }
                }
                value = (int)xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[rowOffset - 1, colOffset].Value;
                myHorzLeft = value;
                endFound = false;
                //Check that connect Excel sheet is formated correctly, azimuth.
                for (int i = 0; !endFound; i++)
                {
                    if (xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[rowOffset - 1, colOffset + i + 1].Value2 != null)
                    {
                        if ((xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[rowOffset - 1, colOffset + i + 1].Value2 == value + 5))
                        {
                            value = (int)xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[rowOffset - 1, colOffset + i + 1].Value2;
                        }
                    }
                    else
                    {
                        myHorzRight = value;
                        endFound = true;
                    }
                }


                int row = 0;
                int col = 0;
                //find closest point on five degree grid
                var rndHz = Math.Round((gsiParser.VerticalAngle / 5)) * 5;
                var rndAz = Math.Round((gsiParser.HorizontalAngle / 5)) * 5;

                var countColumns = ((Math.Abs(myHorzRight) + Math.Abs(myHorzLeft)) / 5) + 1;       //Include zero in the count by adding one.
                var countRows = ((Math.Abs(myVertUpFOV) + Math.Abs(myVertDownFOV)) / 5) + 1;       //Include zero in the count by adding one.

                //find the correct row for the data
                for (int i = 0; i <= countRows; i++)
                {
                    var cellVal = (int)(xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[3 + 2 * i, 1] as Excel.Range).Value;
                    if (cellVal == rndHz)
                    {
                        row = 3 + 2 * i;
                        break;
                    }
                }

                //find the correct column for the data
                for (int j = 1; j <= countColumns; j++)
                {
                    var cellVal = Convert.ToInt32((xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[2, j + 1] as Excel.Range).Value);


                    if (cellVal == rndAz)
                    {
                        col = j + 1;

                        break;
                    }
                }


                xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[row, col] = gsiParser.HorizontalAngle;
                xlWorkBook.Sheets[DATA_COLLECTION_SHEET_NAME].Cells[row + 1, col] = gsiParser.VerticalAngle;
            }

            catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException || ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                myLog(string.Format("serialPort_DataReceived_Auto :: {0}", ex.Message));
            }


        }

        private void listBoxSerialPorts_SelectedIndexChanged(object sender, EventArgs e)
        {
            portName = (string)listBoxSerial.SelectedValue;

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
        
        /*
         * Excel Related
         */
        private void buttonExcelFiles_Click(object sender, EventArgs e)
        {
            //Reset List<Excel.Workbook> workbooks
            workbooks.Clear();
            workbooksStrings.Clear();
            try
            {
                xlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {

            }

            foreach (Excel.Workbook item in xlApp.Workbooks)
            {
                workbooks.Add(item);
                workbooksStrings.Add(item.Name);
            }

            if (workbooks.Count == 0)
            {
                myLog(string.Format("No Excel files found. Open the Excel file first, them open this program."));
            }
            else
            {
                listBoxExcel.SetSelected(0,true);
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

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
                
        private void createDataCollectionSheet(Excel.Worksheet sheet)
        {
            //Set legend cells in upper left corner.
            sheet.Cells[1, 1].Value = "AZ →";
            sheet.Cells[1, 1].Interior.Color = myGrey;
            sheet.Cells[2, 1].Value = "EL ↓";
            sheet.Cells[2, 1].Interior.Color = myGrey;

            //Fill in row 2 with azimuth values. Azimuth values increase by 5 degrees.
            var maxNegativeAz = Convert.ToInt32(-1 * (Convert.ToDecimal(horizontalFOV / 2)));
            var maxPositiveAz = horizontalFOV / 2;
            var tempColumn = 2;                                             //Azimuth legend starts at cell 2.
            for (int i = maxNegativeAz; i <= maxPositiveAz; i += 5, tempColumn++)
            {
                sheet.Cells[2, tempColumn].Value = i;
                sheet.Cells[2, tempColumn].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }

            //Fill in column 1 with elevation values. Row three is where elevation legend starts.
            //For formatting reasons, two cells are merged for each elevation legend value.
            var countUniqueRowValues = (verticalUpFOV + verticalDownFOV) / 5;
            var countElevation = verticalUpFOV;
            for (int i = 0; i <= countUniqueRowValues; i++, countElevation -= 5)
            {
                //Insert value into first cell.
                sheet.Cells[3 + 2 * i, 1] = countElevation;
                //Center text vertically.
                sheet.Cells[3 + 2 * i, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                //Merge first cell with the next cell below.
                sheet.Range[sheet.Cells[3 + 2 * i, 1], sheet.Cells[4 + 2 * i, 1]].Merge();
            }

            /*  
             *  Color the cells.
             *  
             *  The data collection cell blocks are colored in a checkerboard pattern. This helps readability. 
             *  Each entry in the data collection area contains azimuth and elevation information.
             *  Azimuth data is contained in the first cell, elevation is contained in the row below. 
             *  Row 1 column 1 of the data collection area starts at cell[3,2].
             *  Odd columns on odd rows are colorter. Even columns on even rows are colored.
             */
            var countUniqueColumnValues = (horizontalFOV / 5) + 1;          //Include zero in the count by adding one.
            for (int i = 1; i <= countUniqueRowValues + 1; i++)
            {
                for (int j = 1; j <= countUniqueColumnValues; j++)
                {
                    if ((i % 2) != 0)
                    {
                        //Odd row.
                        if (j % 2 != 0)
                        {
                            //Odd column
                            var dataBlock = sheet.Range[sheet.Cells[1 + (2 * i), j + 1], sheet.Cells[1 + (2 * i) + 1, j + 1]];
                            dataBlock.Cells.Interior.Color = myGrey;
                        }

                    }
                    else if (i % 2 == 0)
                    {
                        //Even row
                        if (j % 2 == 0)
                        {
                            //Even column.
                            var dataBlock = sheet.Range[sheet.Cells[1 + (2 * i), j + 1], sheet.Cells[1 + (2 * i) + 1, j + 1]];
                            dataBlock.Cells.Interior.Color = myGrey;
                        }

                    }
                }

            }
        }

        private void createDistortionSheet(Excel.Worksheet sheet)
        {
            //Set legend cells in upper left corner.
            sheet.Cells[1, 1].Value = "AZ →";
            sheet.Cells[1, 1].Interior.Color = myGrey;
            sheet.Cells[2, 1].Value = "EL ↓";
            sheet.Cells[2, 1].Interior.Color = myGrey;

            //Fill in row 2 with azimuth values. Azimuth values increase by 5 degrees.
            var maxNegativeAz = Convert.ToInt32(-1 * (Convert.ToDecimal(horizontalFOV / 2)));
            var maxPositiveAz = horizontalFOV / 2;
            var tempColumn = 2;                                             //Azimuth legend starts at cell 2.
            for (int i = maxNegativeAz; i <= maxPositiveAz; i += 5, tempColumn++)
            {
                sheet.Cells[2, tempColumn].Value = i;
                sheet.Cells[2, tempColumn].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }

            //Fill in column one with elevation values. Row three is where elevation legend starts.           
            var countUniqueRowValues = (verticalUpFOV + verticalDownFOV) / 5;
            var countElevation = verticalUpFOV;
            for (int i = 0; i <= countUniqueRowValues; i++, countElevation -= 5)
            {
                //Insert value into first cell.
                sheet.Cells[3 + i, 1] = countElevation;
            }

            /*  
             *  Color the cells.
             *  
             *  See the function createDataCollectionSheet() for more details.
             */
            var countUniqueColumnValues = (horizontalFOV / 5) + 1;          //Include zero in the count by adding one.
            for (int i = 1; i <= countUniqueRowValues + 1; i++)
            {
                for (int j = 1; j <= countUniqueColumnValues; j++)
                {
                    if ((i % 2) != 0)
                    {
                        //Odd row.
                        if (j % 2 != 0)
                        {
                            //Odd column
                            sheet.Cells[i + 2, j + 1].Interior.Color = myGrey;

                        }

                    }
                    else if (i % 2 == 0)
                    {
                        //Even row
                        if (j % 2 == 0)
                        {
                            //Even column.
                            sheet.Cells[i + 2, j + 1].Interior.Color = myGrey;
                        }

                    }
                }
            }
        }

        private void formatTotalGeometricDistorionSheet(Excel.Workbook book)
        {
            var dcSheet = book.Sheets[DATA_COLLECTION_SHEET_NAME];        //Data collection sheet.
            var tdSheet = book.Sheets[TOTAL_DISTORTION_SHEET_NAME];       //Total distortion sheet.

            //Data collection area starts at row 3, col 2.
            var rowOffset = 3;
            var colOffset = 2;

            List<int> azimuthValues = new List<int>();
            List<int> elevationValues = new List<int>();

            var countUniqueHorizontalValues = (horizontalFOV / 5) + 1;                          //Include zero in the count by adding one.
            var countUniqueElevationValues = ((verticalUpFOV + verticalDownFOV) / 5) + 1;       //Include zero in the count by adding one.

            var azimuth = Convert.ToInt32(-1 * (horizontalFOV / 2));
            for (int i = countUniqueHorizontalValues; i > 0; i--)
            {
                azimuthValues.Add(azimuth);
                azimuth += 5;
            }

            var elevation = Convert.ToInt32(verticalUpFOV);
            for (int i = countUniqueElevationValues; i > 0; i--)
            {
                elevationValues.Add(elevation);
                elevation -= 5;
            }

            //  Insert a formula into each cell to compute the distance between the measured value and the theoretical value.
            //  "=SQRT(((x1 - x2)^2)+((y1 - y2)^2))"
            //
            //  An example of the formula for each cell is
            //      =IF(ISBLANK('Data Collection'!B3),
            //          0,
            //          SQRT((('Data Collection'!B3--25)^2)+(('Data Collection'!B4-10)^2)))
            //
            for (int i = 0; i < countUniqueElevationValues; i++)
            {
                //The variable 'i' tracks the current row in the sheet. 'j' tracks the column.
                for (int j = 0; j < countUniqueHorizontalValues; j++)
                {
                    //For azCell and elCell, the cell name is '[sheet name]'![row,colum]
                    var azCell = string.Format("'{0}'!{1}{2}", DATA_COLLECTION_SHEET_NAME, GetExcelColumnName(j + colOffset), i * 2 + rowOffset);
                    var elCell = string.Format("'{0}'!{1}{2}", DATA_COLLECTION_SHEET_NAME, GetExcelColumnName(j + colOffset), i * 2 + rowOffset + 1);
                    
                    var distFormula = string.Format("SQRT((({0}-{1})^2)+(({2}-{3})^2))", azCell, azimuthValues[j], elCell, elevationValues[i]);

                    //Form the final formula.
                    string finalFormula = @"=IF(ISBLANK(" + azCell + "),0," + distFormula + ")";
                    try
                    {
                        tdSheet.Cells[rowOffset + i, colOffset + j].Formula = @finalFormula;
                    }
                    catch (Exception ex)
                    {
                        myLog(string.Format("formatTotalGeometricDistorion :: {0}", ex.Message));
                    }

                }
            }
        }

        /*
         * Reletive geometry is calculated by comparing one [azimuth, elevation] value to an adjacent [azimuth, elevation] five degrees above, below, left or right.
         * The resulting value is the deviation from a five degree offset, in arc minutes, of the first point to the second.
         */
        private void formatRelativeGeometricDistorionSheet(Excel.Workbook book)
        {
            var dcSheet = book.Sheets[DATA_COLLECTION_SHEET_NAME];          //Data collection sheet.
            var rdSheet = book.Sheets[RELATIVE_DISTORTION_SHEET_NAME];       //Relative distortion sheet.

            //Name short cuts.
            string dc = DATA_COLLECTION_SHEET_NAME;
            string rd = RELATIVE_DISTORTION_SHEET_NAME;

            //Data collection area starts at row 3, col 2.
            var rowOffset = 3;
            var colOffset = 2;

            var countColumns = (horizontalFOV / 5) + 1;                          //Include zero in the count by adding one.
            var countRows = ((verticalUpFOV + verticalDownFOV) / 5) + 1;       //Include zero in the count by adding one.

            //'stringArray' represents the cells in the RELATIVE_DISTORTION_SHEET_NAME sheet.
            //Each string in stringArray will be placed into the excel sheet.
            string[,] stringArray = new string[countRows, countColumns];

            /*             
             * The following nested 'for' statements cycle through each unique index of the 'stringArray'. The 'if' statements check if 
             * the neighbor exists (e.g. not out of bounds of the 2D array). A neighbor is the cell directly above, below, left or right of the current cell.
             * 
             * The Excel formula calculates the distance between neighboring cells using
             *      distance = SQRT( (x2-x1)^2 + (y2-y1)^2 )
             * 
             * The resulting distance (in units Decimal Degrees), is multiplied by 60 (unit is minutes/dec.deg) and subtracted by 300 (unit is minutes). 
             * 300 minutes = 5 dec. deg.
             * 
             * 
             * This is the resultant string for a value with four neighbors
             * =IF(ISBLANK('Data Collection'!C5),,
             *      MAX(IF(ISBLANK('Data Collection'!C3),0,ABS(SQRT(('Data Collection'!C5-'Data Collection'!C3)^2+('Data Collection'!C6-'Data Collection'!C4)^2)*60-300))
             *          ,IF(ISBLANK('Data Collection'!C7),0,ABS(SQRT(('Data Collection'!C5-'Data Collection'!C7)^2+('Data Collection'!C6-'Data Collection'!C8)^2)*60-300))
             *          ,IF(ISBLANK('Data Collection'!B5),0,ABS(SQRT(('Data Collection'!C5-'Data Collection'!B5)^2+('Data Collection'!C6-'Data Collection'!B6)^2)*60-300))
             *          ,IF(ISBLANK('Data Collection'!D5),0,ABS(SQRT(('Data Collection'!C5-'Data Collection'!D5)^2+('Data Collection'!C6-'Data Collection'!D6)^2)*60-300))))
             *          
             *  Excel formulas are messy string concatinations, and I am sorry if you are having to debug here. Hopefully you never read this.
             */
            for (int i = 0; i < countRows; i++)
            {
                for (int j = 0; j < countColumns; j++)
                {
                    string azCell = string.Format("'{0}'!{1}{2}", DATA_COLLECTION_SHEET_NAME, GetExcelColumnName(j + colOffset), i * 2 + rowOffset);
                    string elCell = string.Format("'{0}'!{1}{2}", DATA_COLLECTION_SHEET_NAME, GetExcelColumnName(j + colOffset), i * 2 + rowOffset + 1);

                    string formula = string.Format("=IF(ISBLANK({0}),,MAX(", azCell);

                    if (i - 1 >= 0)
                    {
                        var tempAz = string.Format("'{0}'!{1}{2}", dc, GetExcelColumnName(j + colOffset), i * 2 + rowOffset - 2);
                        var tempEl = string.Format("'{0}'!{1}{2}", dc, GetExcelColumnName(j + colOffset), i * 2 + rowOffset - 1);

                        var isBlank = string.Format("IF(ISBLANK({0}),0,", tempAz);

                        formula = formula + isBlank
                            + "ABS(SQRT((" + azCell + "-" + tempAz + ")^2+(" + elCell + "-" + tempEl + ")^2)*60-300)),";
                    }
                    if (i + 1 < countRows)
                    {
                        var tempAz = string.Format("'{0}'!{1}{2}", dc, GetExcelColumnName(j + colOffset), i * 2 + rowOffset + 2);
                        var tempEl = string.Format("'{0}'!{1}{2}", dc, GetExcelColumnName(j + colOffset), i * 2 + rowOffset + 3);

                        var isBlank = string.Format("IF(ISBLANK({0}),0,", tempAz);

                        formula = formula + isBlank
                           + "ABS(SQRT((" + azCell + "-" + tempAz + ")^2+(" + elCell + "-" + tempEl + ")^2)*60-300)),";
                    }
                    if (j - 1 >= 0)
                    {
                        var tempAz = string.Format("'{0}'!{1}{2}", dc, GetExcelColumnName(j + colOffset - 1), i * 2 + rowOffset);
                        var tempEl = string.Format("'{0}'!{1}{2}", dc, GetExcelColumnName(j + colOffset - 1), i * 2 + rowOffset + 1);

                        var isBlank = string.Format("IF(ISBLANK({0}),0,", tempAz);

                        formula = formula + isBlank
                            + "ABS(SQRT((" + azCell + "-" + tempAz + ")^2+(" + elCell + "-" + tempEl + ")^2)*60-300)),";
                    }
                    if (j + 1 < countColumns)
                    {
                        var tempAz = string.Format("'{0}'!{1}{2}", dc, GetExcelColumnName(j + colOffset + 1), i * 2 + rowOffset);
                        var tempEl = string.Format("'{0}'!{1}{2}", dc, GetExcelColumnName(j + colOffset + 1), i * 2 + rowOffset + 1);

                        var isBlank = string.Format("IF(ISBLANK({0}),0,", tempAz);

                        formula = formula + isBlank
                            + "ABS(SQRT((" + azCell + "-" + tempAz + ")^2+(" + elCell + "-" + tempEl + ")^2)*60-300))";
                    }
                    formula = formula + "))";

                    stringArray[i, j] = formula;
                }

            }

            //Now write the formulas to the Excel sheet.
            for (int i = 0; i < countRows; i++)
            {
                for (int j = 0; j < countColumns; j++)
                {
                    rdSheet.Cells[rowOffset + i, colOffset + j] = stringArray[i, j];
                }
            }

        }

        private void radioButtonManual_CheckedChanged(object sender, EventArgs e)
        {
            autoMode = false;

            groupBoxManualSettings.Visible = true;
            groupBoxManualSettings.Enabled = true;

            checkBoxX.Checked = true;
            checkBoxY.Checked = true;
            checkBoxZ.Checked = true;
            checkBoxAZ.Checked = true;
            checkBoxEL.Checked = true;
            checkBoxDist.Checked = true;
        }

        private void radioButtonAuto_CheckedChanged(object sender, EventArgs e)
        {
            autoMode = true;

            groupBoxManualSettings.Visible = false;
            groupBoxManualSettings.Enabled = false;
        }
        
        /*
         * Data Collection
         */
        private void buttonStartDataCollection_Click(object sender, EventArgs e)
        {
            try
            {
                if (serial.IsOpen && xlWorkBook != null && buttonStartDataCollection.Text == "START WORKING")
                {
                    //The user has requested to start collecting data. All controlls on the winform UI should be disabled except for the "STOP WORKING" button.
                    buttonConnect.Enabled = false;
                    buttonConnect.Visible = false;
                    buttonLockExcel.Enabled = false;
                    buttonLockExcel.Visible = false;

                    //Format the "START WORKING"/"STOP WORKING" data collection button.
                    buttonStartDataCollection.Text = string.Format("STOP WORKING");
                    buttonStartDataCollection.BackColor = myRed;

                    if (autoMode == true)
                    {
                        serial.DataReceived += serialPort_DataReceived_Auto;
                        
                        //Check that connected excel sheet is formatted correctly.

                    }
                    else if (autoMode == false)
                    {
                        serial.DataReceived += serialPort_DataReceived_Manual;
                    }
                }
                else if (buttonStartDataCollection.Text == "STOP WORKING")
                {
                    //Format the "START WORKING"/"STOP WORKING" data collection button.
                    buttonStartDataCollection.Text = string.Format("START WORKING");
                    buttonStartDataCollection.BackColor = myGreen;

                    //Re enable the buttons which were disabled and hidden in the "START WORKING" stage of this button.
                    buttonConnect.Enabled = true;
                    buttonConnect.Visible = true;
                    buttonLockExcel.Enabled = true;
                    buttonLockExcel.Visible = true;

                    if (autoMode == true)
                    {
                        serial.DataReceived -= serialPort_DataReceived_Auto;
                    }
                    else
                    {
                        serial.DataReceived -= serialPort_DataReceived_Manual;
                    }

                }
            }
            catch (Exception ex) when (ex is ArgumentNullException || ex is ArgumentException)
            {
                myLog(string.Format("buttonStartDataCollection_Click() :: {0}", ex.Message));
            }
        }

        private void myLog(string log)
        {
            if(textBoxErrorLog.InvokeRequired)
            {
                var d = new SafeCallDelegate(myLog);
                Invoke(d, new object[] { log });
            }
            else
            {
                if(log == "clear")
                {
                    //Special case, clear the log window.
                    textBoxErrorLog.Clear();                    
                }
                else
                {
                    //Ouput log message to textBox
                    textBoxErrorLog.AppendText(log);
                    if (!log.EndsWith(Environment.NewLine))
                    {
                        textBoxErrorLog.AppendText(Environment.NewLine);
                    }
                }
                
            }            
        }
        
        /*
         * Generate Excel Sheet
         */
        private void buttonGenerate_Click(object sender, EventArgs e)
        {
            myLog("clear");

            //read and format each user input
            if (!string.IsNullOrEmpty(textBoxHorizontalFOV.Text))
            {
                int.TryParse(textBoxHorizontalFOV.Text, out horizontalFOV);
                //Check that the horizontalFOV is a multiple of 10.
                if (horizontalFOV % 10 != 0 || horizontalFOV == 0)
                {
                    myLog(string.Format("Horizontal FOV must be a multiple of 10 and non-zero"));
                    horizontalFOV = 0;
                }
                horizontalFOV = Math.Abs(horizontalFOV);
            }
            if (!string.IsNullOrEmpty(textBoxVerticalUpFOV.Text))
            {
                int.TryParse(textBoxVerticalUpFOV.Text, out verticalUpFOV);
                //Check that the VerticalUpFOV is a multiple of 5.
                if (verticalUpFOV % 5 != 0 || verticalUpFOV == 0)
                {
                    myLog(string.Format("Vertical Up FOV must be a multiple of 5 and non-zero"));
                    verticalUpFOV = 0;
                }
                verticalUpFOV = Math.Abs(verticalUpFOV);

            }
            if (!string.IsNullOrEmpty(textboxVerticalDownFOV.Text))
            {
                int.TryParse(textboxVerticalDownFOV.Text, out verticalDownFOV);
                //Check that the verticalDownFOV is a multiple of 5.
                if (verticalDownFOV % 5 != 0 || verticalDownFOV == 0)
                {
                    myLog(string.Format("Vertical Down FOV must be a multiple of 5 and non-zero"));
                    verticalDownFOV = 0;
                }
                verticalDownFOV = Math.Abs(verticalDownFOV);
            }
            if (!string.IsNullOrEmpty(textBoxDirectory.Text) && Directory.Exists(textBoxDirectory.Text))
            {
                saveFileDir = textBoxDirectory.Text;
            }

            //Check all user inputs. If everything is in order then procede with excel file generation. Note: A value of zero for any FOV is concidered as a missing input.
            if (!string.IsNullOrEmpty(saveFileDir) && Directory.Exists(saveFileDir) && horizontalFOV != 0 && verticalDownFOV != 0 && verticalUpFOV != 0 && eyepoint != "")
            {
                //Format the saveFile string.
                saveFile = string.Format("{0}\\{1}_{2}Up_{3}Down_{4}Horizontal", textBoxDirectory.Text, eyepoint, verticalUpFOV, verticalDownFOV, horizontalFOV);
                myLog(string.Format("Saving file as {0}{1}", Environment.NewLine, saveFile));

                /*
                 * 
                 *  Begin file generation.
                 * 
                */

                Excel.Application xlApp = new Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }

                //Create new work book, create the worksheet variables. 
                var workBook = xlApp.Workbooks.Add();
                var sheet1 = workBook.ActiveSheet as Excel.Worksheet;
                sheet1.Name = RELATIVE_DISTORTION_SHEET_NAME;
                var sheet2 = xlApp.Sheets.Add() as Excel.Worksheet;
                sheet2.Name = TOTAL_DISTORTION_SHEET_NAME;
                var sheet3 = xlApp.Sheets.Add() as Excel.Worksheet;
                sheet3.Name = DATA_COLLECTION_SHEET_NAME;

                try
                {
                    workBook.SaveAs(saveFile);
                    myLog(".");
                    createDataCollectionSheet(sheet3);
                    myLog(".");
                    createDistortionSheet(sheet2);
                    myLog(".");
                    createDistortionSheet(sheet1);

                    formatTotalGeometricDistorionSheet(workBook);
                    formatRelativeGeometricDistorionSheet(workBook);

                    workBook.Save();
                    workBook.Close(0);
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                    myLog(string.Format("Done!"));
                }
                catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException)
                {
                    myLog(string.Format("buttonGenerate_Click :: {0}", ex.Message));
                    if (workBook != null)
                    {
                        workBook.Close(0);
                    }
                }


            }

            //display error messages if the user input is missing
            else
            {
                if (string.IsNullOrEmpty(textBoxHorizontalFOV.Text))
                {
                    myLog(string.Format("Missing Horizontal FOV{0}", Environment.NewLine));
                }
                if (string.IsNullOrEmpty(textBoxVerticalUpFOV.Text))
                {
                    myLog(string.Format("Missing Vertical FOV{0}", Environment.NewLine));
                }
                if (string.IsNullOrEmpty(textboxVerticalDownFOV.Text))
                {
                    myLog(string.Format("Missing Vertical FOV{0}", Environment.NewLine));
                }
                if (string.IsNullOrEmpty(textBoxDirectory.Text) || !Directory.Exists(textBoxDirectory.Text))
                {
                    myLog(string.Format("Directory not found{0}", Environment.NewLine));
                }
                if (eyepoint == "")
                {
                    myLog(string.Format("Eyepoint not set{0}", Environment.NewLine));
                }
            }
        }

        private void radioButtonDEP_CheckedChanged(object sender, EventArgs e)
        {
            eyepoint = "DEP";
        }

        private void radioButtonPEP_CheckedChanged(object sender, EventArgs e)
        {
            eyepoint = "PEP";
        }

        private void radioButtonCPEP_CheckedChanged(object sender, EventArgs e)
        {
            eyepoint = "CPEP";
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            using (var fldrDlg = new FolderBrowserDialog())
            {
                if (fldrDlg.ShowDialog() == DialogResult.OK)
                {
                    saveFileDir = fldrDlg.SelectedPath;
                    textBoxDirectory.Text = saveFileDir;
                }
            }
        }
               
        private void buttonHelp_Click(object sender, EventArgs e)
        {
            myLog("clear");
            myLog(string.Format("**************Serial settings***************{0}(Interface settings on TS02):{0}Baud Rate = {1}{0}Data Bits = {2}{0}Stop Bits = {3}{0}Handshake = {4}{0}Acknowledge = OFF"
                , Environment.NewLine, _baudRate, _dataBits, _stopBits, _handshake));
            myLog("Under Data Settings, set Data Output to Interface");
            myLog(string.Format("************Region Settings************{0}Use DECIMAL DEGREES and METERS", Environment.NewLine));
        }

        private void radioButtonHorizontalManual_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void radioButtonVerticalManual_CheckedChanged(object sender, EventArgs e)
        {
            
        }
    }
}
