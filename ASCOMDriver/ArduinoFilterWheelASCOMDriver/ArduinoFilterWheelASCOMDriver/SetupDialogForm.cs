using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ASCOM.Utilities;
using ASCOM.ArduinoFilterWheel;

namespace ASCOM.ArduinoFilterWheel
{
    [ComVisible(false)]					// Form not registered for COM!
    public partial class SetupDialogForm : Form
    {
        public SetupDialogForm()
        {
            InitializeComponent();
            // Initialise current values of user settings from the ASCOM Profile
            COMBO_COMPort.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());
            if (!COMBO_COMPort.Items.Contains(FilterWheel.comPort) && !String.IsNullOrEmpty(FilterWheel.comPort))
                COMBO_COMPort.Items.Add(FilterWheel.comPort);
            COMBO_COMPort.SelectedItem = FilterWheel.comPort;
            chkTrace.Checked = FilterWheel.traceState;

            TEXTBOX_Filter0.Text = FilterWheel.filter0Name;
            TEXTBOX_Filter1.Text = FilterWheel.filter1Name;
            TEXTBOX_Filter2.Text = FilterWheel.filter2Name;
            TEXTBOX_Filter3.Text = FilterWheel.filter3Name;
            TEXTBOX_Filter4.Text = FilterWheel.filter4Name;

            if (String.IsNullOrEmpty(TEXTBOX_Filter0.Text))
                TEXTBOX_Filter0.Text = "Filter 1";
            if (String.IsNullOrEmpty(TEXTBOX_Filter1.Text))
                TEXTBOX_Filter1.Text = "Filter 2";
            if (String.IsNullOrEmpty(TEXTBOX_Filter2.Text))
                TEXTBOX_Filter2.Text = "Filter 3";
            if (String.IsNullOrEmpty(TEXTBOX_Filter3.Text))
                TEXTBOX_Filter3.Text = "Filter 4";
            if (String.IsNullOrEmpty(TEXTBOX_Filter4.Text))
                TEXTBOX_Filter4.Text = "Filter 5";
        }

        private void cmdOK_Click(object sender, EventArgs e) // OK button event handler
        {
            // Place any validation constraint checks here

            FilterWheel.comPort = COMBO_COMPort.SelectedItem as string; // Update the state variables with results from the dialogue
            FilterWheel.traceState = chkTrace.Checked;
            FilterWheel.filter0Name = TEXTBOX_Filter0.Text;
            FilterWheel.filter1Name = TEXTBOX_Filter1.Text;
            FilterWheel.filter2Name = TEXTBOX_Filter2.Text;
            FilterWheel.filter3Name = TEXTBOX_Filter3.Text;
            FilterWheel.filter4Name = TEXTBOX_Filter4.Text;
        }

        private void cmdCancel_Click(object sender, EventArgs e) // Cancel button event handler
        {
            Close();
        }

        private void BrowseToAscom(object sender, EventArgs e) // Click on ASCOM logo event handler
        {
            try
            {
                System.Diagnostics.Process.Start("http://ascom-standards.org/");
            }
            catch (System.ComponentModel.Win32Exception noBrowser)
            {
                if (noBrowser.ErrorCode == -2147467259)
                    MessageBox.Show(noBrowser.Message);
            }
            catch (System.Exception other)
            {
                MessageBox.Show(other.Message);
            }
        }
    }
}