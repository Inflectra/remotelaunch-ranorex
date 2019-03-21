using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace RanorexAutomationEngine
{
    /// <summary>
    /// Interaction logic for AutomationEngineSettingsPanel.xaml
    /// </summary>
    /// <remarks>
    /// This panel is used to display and automation-engine specific configuration settings
    /// </remarks>
    public partial class AutomationEngineSettingsPanel : UserControl
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public AutomationEngineSettingsPanel()
        {
            InitializeComponent();
            this.LoadSettings();
        }

        /// <summary>
        /// Loads the saved settings
        /// </summary>
        private void LoadSettings()
        {
            //Load the various properties

            this.txtSetting1.Text = Properties.Settings.Default.ResultPath;
            this.chkTraceLogging.IsChecked = Properties.Settings.Default.TraceLogging;
        }

        /// <summary>
        /// Saves the specified settings.
        /// </summary>
        public void SaveSettings()
        {
            //Get the various properties
            Properties.Settings.Default.ResultPath = this.txtSetting1.Text.Trim();
            if (this.chkTraceLogging.IsChecked.HasValue)
            {
                Properties.Settings.Default.TraceLogging = this.chkTraceLogging.IsChecked.Value;
            }

            //Save the properties and reload
            Properties.Settings.Default.Save();
            this.LoadSettings();
        }
    }
}
