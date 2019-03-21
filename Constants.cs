using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RanorexAutomationEngine
{
    /// <summary>
    /// Contains some of the constants and enumerations used by the automation engine
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// The token reported by the automation and used to identify in SpiraTest
        /// </summary>
        public const string AUTOMATION_ENGINE_TOKEN = "RanorexEngine";

        /// <summary>
        /// The version number of the plugin.
        /// </summary>
        public const string AUTOMATION_ENGINE_VERSION = "4.0.1";

        /// <summary>
        /// The name of external automated testing system
        /// </summary>
        public const string EXTERNAL_SYSTEM_NAME = "Ranorex";
    }
}
