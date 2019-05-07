//  ParseGSI
//  Written by: George Marsden III
//  03/29/2019
//  "Leica Geo Serial Interface (GSI) is a generis a general purpose, serial data interface for bi-directional communication
//  between the instrument and the computer.GSI uses a simple command structure to read/write values
//  from/to the sensor.Global and instrument specific Word Indexes (WI) are used to specify various data types."
//
//  This class will parse a GSI formatted string and extract some information. Not all GSI features are supported by this class.
//  The information which is extracted is limited and therefore this class is limited:
//      -All distance data must be in meters.
//      -All spherical measurements must be in decimal degrees.
//      -Only the WIs listed in the private member 'fields' are supported.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TheodoliteReaderUSB
{
    public enum DataFormatGSI { GSI8, GSI16 };
    public enum VerticalAngleMode { ZERO_TO_360, PLUS_MINUS_180 };
    public enum HorizontalAngleMode { ZERO_TO_360, PLUS_MINUS_180 };
    public enum GSIUnits { METER, FEET, GON, DECIMAL_360, SEXAGESIMAL, MIL6400, METER_ONE_TENTH, FEET_ONE_TENTHOUSANTH, METER_ONE_HUNDREDTH}

    class ParseGSI
    {
        //Word Indexes (WI)
        private const string POINTNUMBER = "11";
        private const string HORIZONTAL_ANGLE = "21";
        private const string VERTICAL_ANGLE = "22";
        private const string SLOPE_DISTANCE = "31";
        private const string HORIZONTAL_DISTANCE = "32";
        private const string TARGET_EASTING = "81";
        private const string TARGET_NORTHING = "82";
        private const string TARGET_ELEVATION = "83";

        private static string rawGSIString;

        private static HorizontalAngleMode horizontalAngleMode;
        private static VerticalAngleMode verticalAngleMode;
        private static DataFormatGSI wordFormat;

        private static int pointNumber;
        private static Decimal horizontalAngle;
        private static Decimal verticalAngle;
        private static Decimal slopeDistance;
        private static Decimal horizontalDistance;
        private static Decimal targetEasting;
        private static Decimal targetNorthing;
        private static Decimal targetElevation;

        /// <summary>
        /// Default constructor. Initializes the input string to an empty string. HorizontalAngleMode and VerticalAngleMode default to PLUS_MINUS_180.
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
        /// This method is called everytime the rawGSIString is updated. rawGSIString is updated by set property Raw.
        /// </summary>
        private void update()
        {
            try
            {
                if (rawGSIString.ElementAt(0) == '*')
                {
                    wordFormat = DataFormatGSI.GSI16;
                }
                else
                {
                    wordFormat = DataFormatGSI.GSI8;
                }

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
        /// Use this method to set the raw GSI string. The raw string is then parsed and the resultant values are available.
        /// </summary>
        public string Raw
        {
            set
            {
                rawGSIString = value;
                update();
            }
        }
        public int PointNumber
        {
            get
            {
                return pointNumber;
            }
        }
        public Decimal HorizontalAngle
        {
            get
            {
                return horizontalAngle;
            }
        }
        public Decimal VerticalAngle
        {
            get
            {
                return verticalAngle;
            }
        }
        public Decimal SlopeDistance
        {
            get
            {
                return slopeDistance;
            }
        }
        public Decimal HorizontalDistance
        {
            get
            {
                return horizontalDistance;
            }
        }
        public Decimal TargetEasting
        {
            get
            {
                return targetEasting;
            }
        }
        public Decimal TargetNorthing
        {
            get
            {
                return targetNorthing;
            }
        }
        public Decimal TargetElevation
        {
            get
            {
                return targetElevation;
            }
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
                //Default to max value.
                pointNumber = Int32.MaxValue;
            }

        }
        private void parseHorizontalAngle(string word)
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

            horizontalAngle =  Math.Round(data, 6);
        }
        private void parseVerticalAngle(string word)
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
            else
            {
                throw new Exception("Vertical measurement unit is incorrect. Only use decimal degrees.");
            }

            verticalAngle = Math.Round(data, 6);
            
        }
        private void parseSlopeDistance(string word)
        {
            var units = (Convert.ToInt32(Convert.ToString(word[5])));
            var sign = word[6];
            var data = Convert.ToDecimal(word.Substring(7));

            if(units == (int)GSIUnits.METER)
            {
                //Move decimal.
                if (data != 0)
                {
                    data *= 0.001m;
                }

                slopeDistance = data;
            }
        }
        private void parseHorizontalDistance(string word)
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
        private void parseTargetEasting(string word)
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
        private void parseTargetNorthing(string word)
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
        private void parseTargetElevation(string word)
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

                targetElevation = data;
                horizontalDistance = data;  //I am finding that elevationa dn horizontal distance are always the same.
            }
        }
    }
}
