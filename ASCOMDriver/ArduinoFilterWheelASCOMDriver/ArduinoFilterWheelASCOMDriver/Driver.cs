//tabs=4
// --------------------------------------------------------------------------------
//
// ASCOM FilterWheel driver for ArduinoFilterWheel
//
// Description:	ASCOM driver for a filter wheel assembled by Laurent Brunetto (GAPRA),
//              based on Arduino
//
// Implements:	ASCOM FilterWheel interface version: 2
// Author:		Florian Signoret
//
// Edit Log:
//
// Date			Who	                Vers	Description
// -----------	---	                -----	-------------------------------------------------------
// 26-Aug-2014	Florian Signoret	1.0.0	Initial edit, created from ASCOM driver template
// --------------------------------------------------------------------------------
// 23-Feb-2016	Florian Signoret	1.0.1	Added logs. Wait for feedback after initialization
// --------------------------------------------------------------------------------
//


using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;

using ASCOM;
using ASCOM.Astrometry;
using ASCOM.Astrometry.AstroUtils;
using ASCOM.Utilities;
using ASCOM.DeviceInterface;
using System.Globalization;
using System.Collections;

namespace ASCOM.ArduinoFilterWheel
{
    //
    // Your driver's DeviceID is ASCOM.ArduinoFilterWheel.FilterWheel
    //
    // The Guid attribute sets the CLSID for ASCOM.ArduinoFilterWheel.FilterWheel
    // The ClassInterface/None addribute prevents an empty interface called
    // _ArduinoFilterWheel from being created and used as the [default] interface
    //
    //

    /// <summary>
    /// ASCOM FilterWheel Driver for ArduinoFilterWheel.
    /// </summary>
    [Guid("41779da8-218e-449e-bef5-e66450a2c5e0")]
    [ClassInterface(ClassInterfaceType.None)]
    public class FilterWheel : IFilterWheelV2
    {
        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal static string driverID = "ASCOM.ArduinoFilterWheel.FilterWheel";

        /// <summary>
        /// Driver description that displays in the ASCOM Chooser.
        /// </summary>
        private static string driverDescription = "Arduino FilterWheel";

        internal static string comPortProfileName = "COM Port"; // Constants used for Profile persistence
        internal static string comPortDefault = "COM1";
        internal static string traceStateProfileName = "Trace Level";
        internal static string traceStateDefault = "false";
        internal static string filter0ProfileName = "Filter0Name";
        internal static string filter1ProfileName = "Filter1Name";
        internal static string filter2ProfileName = "Filter2Name";
        internal static string filter3ProfileName = "Filter3Name";
        internal static string filter4ProfileName = "Filter4Name";

        internal static string comPort; // Variables to hold the current device configuration
        internal static bool traceState;
        internal static string filter0Name;
        internal static string filter1Name;
        internal static string filter2Name;
        internal static string filter3Name;
        internal static string filter4Name;

        /// <summary>
        /// Private variable to hold an ASCOM Utilities object
        /// </summary>
        private Util utilities;

        /// <summary>
        /// Private variable to hold an ASCOM AstroUtilities object to provide the Range method
        /// </summary>
        private AstroUtils astroUtilities;

        /// <summary>
        /// Private variable to hold the trace logger object (creates a diagnostic log file with information that you specify)
        /// </summary>
        private TraceLogger tl;

        /// <summary>
        /// Private variable to hold the serial communication object
        /// </summary>
        private ASCOM.Utilities.Serial serialPort;

        /// <summary>
        /// Constant that specifies the speed for the serial port: 9600 bps
        /// </summary>
        private const SerialSpeed SERIAL_PORT_SPEED = SerialSpeed.ps9600;

        /// <summary>
        /// Constant that specifies the command delimiter character
        /// </summary>
        private const char COMMAND_DELIMITER_CHAR = '#';

        /// <summary>
        /// Constant that specifies the name of the GETFILTER command
        /// </summary>
        private const string GETFILTER_COMMAND = "GETFILTER";

        /// <summary>
        /// Constant that specifies the name of the SETFILTER command
        /// </summary>
        private const string SETFILTER_COMMAND = "SETFILTER";

        /// <summary>
        /// Initializes a new instance of the <see cref="ArduinoFilterWheel"/> class.
        /// Must be public for COM registration.
        /// </summary>
        public FilterWheel()
        {
            ReadProfile(); // Read device configuration from the ASCOM Profile store

            tl = new TraceLogger("", "ArduinoFilterWheel");
            tl.Enabled = traceState;
            tl.LogMessage("FilterWheel", "Starting initialisation");

            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro utilities object

            tl.LogMessage("FilterWheel", "Completed initialisation");
        }


        //
        // PUBLIC COM INTERFACE IFilterWheelV2 IMPLEMENTATION
        //

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialog form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public void SetupDialog()
        {
            // consider only showing the setup dialog if not connected
            // or call a different dialog if connected
            if (IsConnected)
                System.Windows.Forms.MessageBox.Show("Already connected, just press OK");

            using (SetupDialogForm F = new SetupDialogForm())
            {
                var result = F.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }

        public ArrayList SupportedActions
        {
            get
            {
                tl.LogMessage("SupportedActions Get", "Returning empty arraylist");
                return new ArrayList();
            }
        }

        public string Action(string actionName, string actionParameters)
        {
            throw new ASCOM.ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
        }

        public void CommandBlind(string command, bool raw)
        {
            CheckConnected("CommandBlind");
            throw new ASCOM.MethodNotImplementedException("CommandBlind");
        }

        public bool CommandBool(string command, bool raw)
        {
            CheckConnected("CommandBool");
            throw new ASCOM.MethodNotImplementedException("CommandBool");
        }

        public string CommandString(string command, bool raw)
        {
            CheckConnected("CommandString");
            // it's a good idea to put all the low level communication with the device here,
            // then all communication calls this function
            // you need something to ensure that only one command is in progress at a time

            throw new ASCOM.MethodNotImplementedException("CommandString");
        }

        public void Dispose()
        {
            // Clean up the tracelogger and util objects
            tl.Enabled = false;
            tl.Dispose();
            tl = null;
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;

            if(serialPort != null)
            {
                serialPort.Connected = false;
            }
        }

        public bool Connected
        {
            get
            {
                tl.LogMessage("Connected Get", IsConnected.ToString());
                return IsConnected;
            }
            set
            {
                tl.LogMessage("Connected Set", value.ToString());
                if (value == IsConnected)
                    return;

                if (value)
                {
                    tl.LogMessage("Connected Set", "Connecting to port " + comPort);

                    try
                    {
                        lock (lockIsBusy)
                        {
                            isBusy = true;
                        }

                        // try to parse the COM port number
                        int comPortNumber = 0;
                        string comPortNumberString = String.IsNullOrEmpty(comPort) ? String.Empty : comPort.ToUpper().Replace("COM", "");
                        if (!Int32.TryParse(comPortNumberString, out comPortNumber))
                        {
                            throw new DriverException(String.Format("Could not parse the provided COM Port: [{0}]", comPort));
                        }

                        // Create the serial communication object
                        serialPort = new Serial()
                        {
                            Port = comPortNumber,
                            Speed = SERIAL_PORT_SPEED,
                            Connected = true
                        };

                        tl.LogMessage("Connected Set", "Serial communication object created");

                        tl.LogMessage("Serial port", "Waiting for response after connection");
                        string response = serialPort.ReceiveTerminated(COMMAND_DELIMITER_CHAR.ToString());
                        tl.LogMessage("Serial port", "Received response" + response);
                    }
                    finally
                    {
                        lock (lockIsBusy)
                        {
                            isBusy = false;
                        }
                    }
                }
                else
                {
                    tl.LogMessage("Connected Set", "Disconnecting from port " + comPort);
                    
                    if(serialPort != null)
                    {
                        serialPort.Connected = false;
                        serialPort.Dispose();
                        serialPort = null;
                    }

                    tl.LogMessage("Connected Set", "Disconnected from port " + comPort);
                }
            }
        }

        public string Description
        {
            get
            {
                tl.LogMessage("Description Get", driverDescription);
                return driverDescription;
            }
        }

        public string DriverInfo
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverInfo = "Driver for Arduino FilterWheel. Design: Laurent Brunetto. Author: Florian Signoret. Version: "
                    + String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                tl.LogMessage("InterfaceVersion Get", "2");
                return Convert.ToInt16("2");
            }
        }

        public string Name
        {
            get
            {
                tl.LogMessage("Name Get", driverDescription);
                return driverDescription;
            }
        }

        #endregion

        #region IFilerWheel Implementation
        private int[] fwOffsets = new int[5] { 0, 0, 0, 0, 0 }; //class level variable to hold focus offsets
        private string[] fwNames = new string[5] { "Filter 1", "Filter 2", "Filter 3", "Filter 4", "Filter 5" }; //class level variable to hold the filter names

        public int[] FocusOffsets
        {
            get
            {
                foreach (int fwOffset in fwOffsets) // Write filter offsets to the log
                {
                    tl.LogMessage("FocusOffsets Get", fwOffset.ToString());
                }

                return fwOffsets;
            }
        }

        public string[] Names
        {
            get
            {
                foreach (string fwName in fwNames) // Write filter names to the log
                {
                    tl.LogMessage("Names Get", fwName);
                }

                return fwNames;
            }
        }

        private bool isBusy;
        private object lockIsBusy = new Object();

        public short Position
        {
            get
            {
                tl.LogMessage("Position Get", "function called");

                if (isBusy)
                    return -1;
                
                short currentPosition = GetCurrentPosition();
                tl.LogMessage("Position Get", currentPosition.ToString());
                return currentPosition;
            }
            set
            {
                tl.LogMessage("Position Set", value.ToString());
                if ((value < 0) | (value > fwNames.Length - 1))
                {
                    tl.LogMessage("", "Throwing InvalidValueException - Position: " + value.ToString() + ", Range: 0 to " + (fwNames.Length - 1).ToString());
                    throw new InvalidValueException("Position", value.ToString(), "0 to " + (fwNames.Length - 1).ToString());
                }

                MoveFilterWheel(value);

                tl.LogMessage("Position Set", "terminated");
            }
        }

        #endregion

        #region Private properties and methods
        // here are some useful properties and methods that can be used as required
        // to help with driver development

        #region ASCOM Registration

        // Register or unregister driver for ASCOM. This is harmless if already
        // registered or unregistered. 
        //
        /// <summary>
        /// Register or unregister the driver with the ASCOM Platform.
        /// This is harmless if the driver is already registered/unregistered.
        /// </summary>
        /// <param name="bRegister">If <c>true</c>, registers the driver, otherwise unregisters it.</param>
        private static void RegUnregASCOM(bool bRegister)
        {
            using (var P = new ASCOM.Utilities.Profile())
            {
                P.DeviceType = "FilterWheel";
                if (bRegister)
                {
                    P.Register(driverID, driverDescription);
                }
                else
                {
                    P.Unregister(driverID);
                }
            }
        }

        /// <summary>
        /// This function registers the driver with the ASCOM Chooser and
        /// is called automatically whenever this class is registered for COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is successfully built.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During setup, when the installer registers the assembly for COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually register a driver with ASCOM.
        /// </remarks>
        [ComRegisterFunction]
        public static void RegisterASCOM(Type t)
        {
            RegUnregASCOM(true);
        }

        /// <summary>
        /// This function unregisters the driver from the ASCOM Chooser and
        /// is called automatically whenever this class is unregistered from COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is cleaned or prior to rebuilding.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During uninstall, when the installer unregisters the assembly from COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually unregister a driver from ASCOM.
        /// </remarks>
        [ComUnregisterFunction]
        public static void UnregisterASCOM(Type t)
        {
            RegUnregASCOM(false);
        }

        #endregion

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private bool IsConnected
        {
            get
            {
                tl.LogMessage("IsConnected", serialPort != null && serialPort.Connected ? "connected" : "disconnected");
                // Check that the driver hardware connection exists and is connected to the hardware
                return serialPort != null && serialPort.Connected;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private void CheckConnected(string message)
        {
            if (!IsConnected)
            {
                throw new ASCOM.NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal void ReadProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "FilterWheel";
                traceState = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, string.Empty, traceStateDefault));
                comPort = driverProfile.GetValue(driverID, comPortProfileName, string.Empty, comPortDefault);

                fwNames[0] = driverProfile.GetValue(driverID, filter0ProfileName, String.Empty, "Filter 1");
                fwNames[1] = driverProfile.GetValue(driverID, filter1ProfileName, String.Empty, "Filter 2");
                fwNames[2] = driverProfile.GetValue(driverID, filter2ProfileName, String.Empty, "Filter 3");
                fwNames[3] = driverProfile.GetValue(driverID, filter3ProfileName, String.Empty, "Filter 4");
                fwNames[4] = driverProfile.GetValue(driverID, filter4ProfileName, String.Empty, "Filter 5");
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "FilterWheel";
                driverProfile.WriteValue(driverID, traceStateProfileName, traceState.ToString());
                driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString());
                driverProfile.WriteValue(driverID, filter0ProfileName, filter0Name);
                driverProfile.WriteValue(driverID, filter1ProfileName, filter1Name);
                driverProfile.WriteValue(driverID, filter2ProfileName, filter2Name);
                driverProfile.WriteValue(driverID, filter3ProfileName, filter3Name);
                driverProfile.WriteValue(driverID, filter4ProfileName, filter4Name);
            }
        }

        #endregion

        #region Serial Port Communication
        
        private short GetCurrentPosition()
        {
            tl.LogMessage("Serial port", String.Format("Sending {0}{1}", GETFILTER_COMMAND, COMMAND_DELIMITER_CHAR));
            serialPort.Transmit(String.Format("{0}{1}", GETFILTER_COMMAND, COMMAND_DELIMITER_CHAR));
            tl.LogMessage("Serial port", "Waiting for response");
            string response = serialPort.ReceiveTerminated(COMMAND_DELIMITER_CHAR.ToString());
            tl.LogMessage("Serial port", "Received response" + response);
            string responseContent = response.Remove(response.IndexOf(COMMAND_DELIMITER_CHAR));
            return Int16.Parse(responseContent);
        }

        private void MoveFilterWheel(short position)
        {
            try
            {
                lock (lockIsBusy)
                {
                    isBusy = true;
                }
                tl.LogMessage("Serial port", String.Format("Sending {0}{1}{2}", SETFILTER_COMMAND, position.ToString(),
                    COMMAND_DELIMITER_CHAR));
                serialPort.Transmit(String.Format("{0}{1}{2}", SETFILTER_COMMAND, position.ToString(),
                    COMMAND_DELIMITER_CHAR));
                tl.LogMessage("Serial port", "Waiting for response");
                serialPort.ReceiveTerminated(COMMAND_DELIMITER_CHAR.ToString());
                tl.LogMessage("Serial port", "Received response");
            }
            finally
            {
                lock (lockIsBusy)
                {
                    isBusy = false;
                }
            }
        }

        #endregion
    }
}
