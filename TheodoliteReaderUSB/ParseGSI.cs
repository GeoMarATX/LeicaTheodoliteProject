//  ParseGSI
//  Written by: George Marsden III
//  03/29/2019
//
//  "Leica Geo Serial Interface (GSI) is a generis a general purpose, 
//  serial data interface for bi-directional communication between the 
//  instrument and the computer. GSI uses a simple command structure to 
//  read/write values from/to the sensor. Global and instrument specific 
//  Word Indexes (WI) are used to specify various data types."
//  Excerpt from:
//  https://w3.leica-geosystems.com/media/new/product_solution/gsi_manual.pdf
//
//  This class will parse a GSI formatted string and then extract some information. 
//  The extracted information can be retreived by the various 'get' methods.
//  
//  Not all GSI features are supported by this class. Only METER and DECIMAL_360
//  units are currently supported.
//
// EXAMPLE GSI8:
//      "110001+0000A110 81..00+00005387 82..00-00000992"

using System;
using System.Linq;

namespace TheodoliteReaderUSB
{
    public enum GSIDataFormat { GSI8, GSI16 };
    public enum VerticalAngleMode { ZERO_TO_360, PLUS_MINUS_180 };
    public enum HorizontalAngleMode { ZERO_TO_360, PLUS_MINUS_180 };
    public enum GSIUnits { METER, FEET, GON, DECIMAL_360, SEXAGESIMAL, MIL6400, METER_ONE_TENTH, FEET_ONE_TENTHOUSANTH, METER_ONE_HUNDREDTH }    //  Only METER and DECIMAL_360 units are currently supported.

    class ParseGSI
    {
        //  Word Indexes (WI)
        private const string POINTNUMBER = "11";
        private const string HORIZONTAL_ANGLE = "21";
        private const string VERTICAL_ANGLE = "22";
        private const string SLOPE_DISTANCE = "31";
        private const string HORIZONTAL_DISTANCE = "32";
        private const string TARGET_EASTING = "81";
        private const string TARGET_NORTHING = "82";
        private const string TARGET_ELEVATION = "83";

        private static string rawGSIString;

        //  Measurement modes
        private static HorizontalAngleMode horizontalAngleMode;
        private static VerticalAngleMode verticalAngleMode;
        private static GSIDataFormat wordFormat;

        //  Extracted GSI data
        private static int pointNumber;
        private static Decimal horizontalAngle;
        private static Decimal verticalAngle;
        private static Decimal slopeDistance;
        private static Decimal horizontalDistance;
        private static Decimal targetEasting;
        private static Decimal targetNorthing;
        private static Decimal targetElevation;

        /// <summary>
        /// Default constructor.
        /// </summary>
        public ParseGSI()
        {
            rawGSIString = "";
            horizontalAngleMode = HorizontalAngleMode.PLUS_MINUS_180;
            verticalAngleMode = VerticalAngleMode.PLUS_MINUS_180;
        }
        /// <summary>
        /// Constructor. User supplies the HorizontalAngleMode.
        /// </summary>
        /// <param name="hzMode"></param>
        public ParseGSI(HorizontalAngleMode hzMode) : base()
        {
            horizontalAngleMode = hzMode;
        }
        /// <summary>
        /// Constructor. User supplies the VerticalAngleMode.
        /// </summary>
        /// <param name="vrMode"></param>
        public ParseGSI(VerticalAngleMode vrMode) : base()
        {
            verticalAngleMode = vrMode;
        }
        /// <summary>
        /// Constructor. User supplies the HorizontalAngleMode and the VerticalAngleMode.
        /// </summary>
        /// <param name="hzMode"></param>
        /// <param name="vrMode"></param>
        public ParseGSI(HorizontalAngleMode hzMode, VerticalAngleMode vrMode) : base()
        {
            horizontalAngleMode = hzMode;
            verticalAngleMode = vrMode;
        }

        /// <summary>
        /// This method is called everytime rawGSIString is set.
        /// </summary>
        private void update()
        {
            try
            {
                wordFormat = (rawGSIString.ElementAt(0) == '*') ? GSIDataFormat.GSI16 :GSIDataFormat.GSI8;
                
                rawGSIString = rawGSIString.ToString().TrimEnd('\r', '\n');
                rawGSIString = rawGSIString.ToString().TrimEnd(' ');
                string[] words = rawGSIString.Split(' ');

                foreach (string word in words)
                {
                    string c = word.Substring(0, 2);
                    switch (c)
                    {
                        case POINTNUMBER:
                            parsePointNumber(word);
                            break;
                        case HORIZONTAL_ANGLE:
                            parseHorizontalAngle(word);
                            break;
                        case VERTICAL_ANGLE:
                            parseVerticalAngle(word);
                            break;
                        case SLOPE_DISTANCE:
                            parseSlopeDistance(word);
                            break;
                        case HORIZONTAL_DISTANCE:
                            parseHorizontalDistance(word);
                            break;
                        case TARGET_EASTING:
                            parseTargetEasting(word);
                            break;
                        case TARGET_NORTHING:
                            parseTargetNorthing(word);
                            break;
                        case TARGET_ELEVATION:
                            parseTargetElevation(word);
                            break;
                        default:
                            break;
                    }
                }
            }
            catch
            {

            }

        }

        /// <summary>
        /// Use this method to set the raw GSI string. The raw string is then parsed by a call to update().
        /// </summary>
        public string Raw
        {
            set
            {
                rawGSIString = value;
                update();
            }
        }
    
        //  Retreive parsed data.
        public int PointNumber
        {
            get => pointNumber;
        }
        public Decimal HorizontalAngle
        {
            get => horizontalAngle;
        }
        public Decimal VerticalAngle
        {
            get => verticalAngle;
        }
        public Decimal SlopeDistance
        {
            get => slopeDistance;
        }
        public Decimal HorizontalDistance
        {
            get => horizontalDistance;
        }
        public Decimal TargetEasting
        {
            get => targetEasting;
        }
        public Decimal TargetNorthing
        {
            get => targetNorthing;
        }
        public Decimal TargetElevation
        {
            get => targetElevation;
        }
        
        private void parsePointNumber(string word)
        {
            try
            {
                pointNumber = Convert.ToInt32(word.Substring(word.IndexOf("+")));
            }
            catch
            {               
            }
            finally
            {
                //Default to max value to indicate that an error has occured.
                pointNumber = Int32.MaxValue;
            }

        }

        private void parseHorizontalAngle(string word)
        {
            try
            {
                var units = (Convert.ToInt32(Convert.ToString(word[5])));
                var sign = word[6];
                var data = Convert.ToDecimal(word.Substring(7));

                if (data != 0)
                {
                    data *= 0.00001m;
                }

                if (units == (int)GSIUnits.DECIMAL_360)
                {
                    switch (horizontalAngleMode)
                    {
                        case HorizontalAngleMode.PLUS_MINUS_180:
                            //convert from 0 to 360, to +- 180
                            if (data >= 180)
                            {
                                data -= 360;
                            }
                            //'+' indicates that theodolite is set to increase to the right.
                            if (sign == '-')
                            {
                                data *= -1;
                            }
                            break;
                        case HorizontalAngleMode.ZERO_TO_360:

                            break;

                        default:
                            break;
                    }

                }
                horizontalAngle = Math.Round(data, 6);
            }
            catch
            {
            }
            finally
            {
                //Default to max value to indicate that an error has occured.
                horizontalAngle = Int32.MaxValue;
            }
        }

        private void parseVerticalAngle(string word)
        {
            try
            {
                var units = (Convert.ToInt32(Convert.ToString(word[5])));
                var sign = word[6];
                var data = Convert.ToDecimal(word.Substring(7));

                if (data != 0.0m)
                {
                    data *= 0.00001m;
                }

                if (units == (int)GSIUnits.DECIMAL_360)
                {
                    switch (verticalAngleMode)
                    {
                        case VerticalAngleMode.PLUS_MINUS_180:
                            {
                                if (data >= 180 && data < 360)
                                {
                                    data -= 270;
                                }
                                else if (data < 180 && data >= 0)
                                {
                                    data -= 90;
                                    data *= -1;
                                }
                            }
                            break;
                        case VerticalAngleMode.ZERO_TO_360:
                            {

                            }
                            break;

                    }
                }
                verticalAngle = Math.Round(data, 6);
            }
            catch
            {
            }
            finally
            {
                //Default to max value to indicate that an error has occured.
                verticalAngle = Int32.MaxValue;
            }

        }

        private void parseSlopeDistance(string word)
        {
            try
            {
                var units = (Convert.ToInt32(Convert.ToString(word[5])));
                var sign = word[6];
                var data = Convert.ToDecimal(word.Substring(7));

                if (units == (int)GSIUnits.METER)
                {
                    //Move decimal.
                    if (data != 0)
                    {
                        data *= 0.001m;
                    }

                    slopeDistance = data;
                }
            }
            catch
            {
            }
            finally
            {
                //Default to max value to indicate that an error has occured.
                slopeDistance = Int32.MaxValue;
            }
        }

        private void parseHorizontalDistance(string word)
        {
            try
            {
                var units = (Convert.ToInt32(Convert.ToString(word[5])));
                var sign = word[6];
                var data = Convert.ToDecimal(word.Substring(7));

                if (units == (int)GSIUnits.METER)
                {
                    //Move decimal.
                    if (data != 0)
                    {
                        data *= 0.001m;
                    }
                    if (sign == '-')
                    {
                        data *= -1;
                    }
                    horizontalDistance = data;
                }
            }
            catch
            {
            }
            finally
            {
                //Default to max value to indicate that an error has occured.
                horizontalDistance = Int32.MaxValue;
            }

        }

        private void parseTargetEasting(string word)
        {
            try
            {
                var units = (Convert.ToInt32(Convert.ToString(word[5])));
                var sign = word[6];
                var data = Convert.ToDecimal(word.Substring(7));

                if (units == (int)GSIUnits.METER)
                {
                    //Move decimal.
                    if (data != 0)
                    {
                        data *= 0.001m;
                    }

                    if (sign == '-')
                    {
                        data *= -1;
                    }

                    targetEasting = data;
                }
            }
            catch
            {
            }
            finally
            {
                //Default to max value to indicate that an error has occured.
                targetEasting = Int32.MaxValue;
            }

        }

        private void parseTargetNorthing(string word)
        {
            try
            {
                var units = (Convert.ToInt32(Convert.ToString(word[5])));
                var sign = word[6];
                var data = Convert.ToDecimal(word.Substring(7));

                if (units == (int)GSIUnits.METER)
                {
                    //Move decimal.
                    if (data != 0)
                    {
                        data *= 0.001m;
                    }

                    if (sign == '-')
                    {
                        data *= -1;
                    }

                    targetNorthing = data;
                }
            }
            catch
            {
            }
            finally
            {
                //Default to max value to indicate that an error has occured.
                targetNorthing = Int32.MaxValue;
            }
        }

        private void parseTargetElevation(string word)
        {
            try
            {
                var units = (Convert.ToInt32(Convert.ToString(word[5])));
                var sign = word[6];
                var data = Convert.ToDecimal(word.Substring(7));

                if (units == (int)GSIUnits.METER)
                {
                    //Move decimal.
                    if (data != 0)
                    {
                        data *= 0.001m;
                    }

                    if (sign == '-')
                    {
                        data *= -1;
                    }
                    //Target elevation and horizontal distance are the same measurement.
                    targetElevation = data;
                    horizontalDistance = data;  
                }
            }
            catch
            {
            }
            finally
            {
                //Default to max value to indicate that an error has occured.
                targetElevation = Int32.MaxValue;
                horizontalDistance = Int32.MaxValue;
            }

        }
    }
}
