using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Inflectra.RemoteLaunch.Interfaces;
using Inflectra.RemoteLaunch.Interfaces.DataObjects;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml;

namespace RanorexAutomationEngine
{
    /// <summary>
    /// Sample data-synchronization provider that synchronizes incidents between SpiraTest/Plan/Team and an external system
    /// </summary>

    /// <summary>
    /// Sample test automation engine plugin that implements the IAutomationEngine class.
    /// This class is instantiated by the RemoteLaunch application
    /// </summary>
    /// <remarks>
    /// The AutomationEngine class provides some of the generic functionality
    /// </remarks>
    public class RanorexAutomationEngine : AutomationEngine, IAutomationEngine4
    {
        private const string CLASS_NAME = "RanorexAutomationEngine";

        /// <summary>
        /// Constructor
        /// </summary>
        public RanorexAutomationEngine()
        {
            //Set status to OK
            base.status = EngineStatus.OK;
        }


        /// <summary>
        /// Returns the author of the test automation engine
        /// </summary>
        public override string ExtensionAuthor
        {
            get
            {
                return "step2IT GmbH";
            }
        }

        /// <summary>
        /// The unique GUID that defines this automation engine
        /// </summary>
        public override Guid ExtensionID
        {
            get
            {
                return new Guid("{714A64BE-78F3-4D17-89BC-984B95D58E6A}");

            }
        }

        /// <summary>
        /// Returns the display name of the automation engine
        /// </summary>
        public override string ExtensionName
        {
            get
            {
                return "Ranorex Automation Engine";
            }
        }

        /// <summary>
        /// Returns the unique token that identifies this automation engine to SpiraTest
        /// </summary>
        public override string ExtensionToken
        {
            get
            {
                return Constants.AUTOMATION_ENGINE_TOKEN;
            }
        }

        /// <summary>
        /// Returns the version number of this extension
        /// </summary>
        public override string ExtensionVersion
        {
            get
            {
                return Constants.AUTOMATION_ENGINE_VERSION;
            }
        }

        /// <summary>
        /// Adds a custom settings panel for allowing the user to set any engine-specific configuration values
        /// </summary>
        /// <remarks>
        /// 1) If you don't have any engine-specific settings, just comment out the entire Property
        /// 2) The SettingPanel needs to be implemented as a WPF XAML UserControl
        /// </remarks>
        public override System.Windows.UIElement SettingsPanel
        {
            get
            {
                return new AutomationEngineSettingsPanel();
            }
            set
            {
                AutomationEngineSettingsPanel settingsPanel = (AutomationEngineSettingsPanel)value;
                settingsPanel.SaveSettings();
            }
        }

        /* Leave this function un-implemented */
        public override AutomatedTestRun StartExecution(AutomatedTestRun automatedTestRun)
        {
            //Not used since we implement the V4 API instead
            throw new NotImplementedException();
        }

        /// <summary>
        /// This is the main method that is used to start automated test execution
        /// </summary>
        /// <param name="automatedTestRun">The automated test run object</param>
        /// <param name="projectId">The id of the project</param>
        /// <returns>Either the populated test run or an exception</returns>
        public AutomatedTestRun4 StartExecution(AutomatedTestRun4 automatedTestRun, int projectId)
        {
            //Set status to OK
            base.status = EngineStatus.OK;
            string externalTestDetailedResults = "";
            string sCompletePath = "";
           
            try
            {
                if (Properties.Settings.Default.TraceLogging)
                {
                    LogEvent("Starting test execution", EventLogEntryType.Information);
                }
                DateTime startDate = DateTime.Now;

                //See if we have any parameters we need to pass to the automation engine
                Dictionary<string, string> parameters = new Dictionary<string, string>();
                if (automatedTestRun.Parameters == null)
                {
                    if (Properties.Settings.Default.TraceLogging)
                    {
                        LogEvent("Test Run has no parameters", EventLogEntryType.Information);
                    }
                }
                else
                {
                    if (Properties.Settings.Default.TraceLogging)
                    {
                        LogEvent("Test Run has parameters", EventLogEntryType.Information);
                    }

                    foreach (TestRunParameter testRunParameter in automatedTestRun.Parameters)
                    {
                        string parameterName = testRunParameter.Name.Trim();
                        if (!parameters.ContainsKey(parameterName))
                        {
                            //Make sure the parameters are lower case
                            if (Properties.Settings.Default.TraceLogging)
                            {
                                LogEvent("Adding test run parameter " + parameterName + " = " + testRunParameter.Value, EventLogEntryType.Information);
                            }
                            parameters.Add(parameterName, testRunParameter.Value);
                        }
                    }
                }

                string runnerTestName = "Unknown";

                //See if we have an attached or linked test script
                if (automatedTestRun.Type == AutomatedTestRun4.AttachmentType.URL)
                {
                    //The "URL" of the test is actually the full file path of the file that contains the test script
                    //Some automation engines need additional parameters which can be provided by allowing the test script filename
                    //to consist of multiple elements separated by a specific character.
                    //Conventionally, most engines use the pipe (|) character to delimit the different elements
                    string[] elements = automatedTestRun.FilenameOrUrl.Split('|');
                    
                    //To make it easier, we have certain shortcuts that can be used in the path
                    //This allows the same test to be run on different machines with different physical folder layouts
                    string path = elements[0];
                    runnerTestName = Path.GetFileNameWithoutExtension(path);
                    string args = "";
                    if (elements.Length > 1)
                    {
                        args = elements[1];
                    }
                    path = path.Replace("[MyDocuments]", Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments));
                    path = path.Replace("[CommonDocuments]", Environment.GetFolderPath(System.Environment.SpecialFolder.CommonDocuments));
                    path = path.Replace("[DesktopDirectory]", Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory));
                    path = path.Replace("[ProgramFiles]", Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles));
                    path = path.Replace("[ProgramFilesX86]", Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFilesX86));

                    //First make sure that the file exists
                    if (File.Exists(path))
                    {
                        if (Properties.Settings.Default.TraceLogging)
                        {
                            LogEvent("Executing " + Constants.EXTERNAL_SYSTEM_NAME + " test located at " + path, EventLogEntryType.Information);
                        }

                        //Get the working dir
                        string workingDir = Path.GetDirectoryName(path);

                        //Create the folder that will be used to store the output file
                        string sResultFile= "TS" + automatedTestRun.TestSetId + "_TC" + automatedTestRun.TestCaseId;
                        string directory = Properties.Settings.Default.ResultPath + "\\" + GetTimestamp(DateTime.Now) + "_"+sResultFile+"\\";
                        sCompletePath=directory + sResultFile + ".rxlog";
                        CreateDirectory(new DirectoryInfo(directory)); 

                        // start the process and wait for exit
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        StringBuilder builder = new StringBuilder("/rf:\"" + sCompletePath+"\"");

                        //Convert the parameters into Ranorex format
                        //  /param:Parameter1="New Value"
                        foreach (KeyValuePair<string,string> parameter in parameters)
                        {
                            builder.Append(" /param:" + parameter.Key + "=\"" + parameter.Value + "\"");
                        }

                        //Add on any user-specified arguments
                        if (!String.IsNullOrWhiteSpace(args))
                        {
                            builder.Append(" " + args);
                        }

                        startInfo.Arguments = builder.ToString();
                        startInfo.FileName = path;
                        startInfo.WorkingDirectory = workingDir;
                        startInfo.RedirectStandardOutput = true;
                        startInfo.UseShellExecute = false;
                        Process p = Process.Start(startInfo);
                        externalTestDetailedResults = String.Format("Executing: {0} in '{1}' with arguments '{2}'\n", startInfo.FileName, startInfo.WorkingDirectory, startInfo.Arguments);
                        externalTestDetailedResults += p.StandardOutput.ReadToEnd();
                        p.WaitForExit();
                        p.Close();

                    }
                    else
                    {
                        throw new FileNotFoundException("Unable to find a " + Constants.EXTERNAL_SYSTEM_NAME + " test at " + path);
                    }
                }
                else
                {
                    //We have an embedded script which we need to send to the test execution engine
                    //If the automation engine doesn't support embedded/attached scripts, throw the following exception:
                    throw new InvalidOperationException("The " + Constants.EXTERNAL_SYSTEM_NAME + " automation engine only supports linked test scripts");
                }

                //Capture the time that it took to run the test
                DateTime endDate = DateTime.Now;

                //Now extract the test results
                //Ranorex saves the XML data in a .rxlog.data file
                XmlDocument doc = new XmlDocument();             
                doc.Load(sCompletePath + ".data");

                // Select the first book written by an author whose last name is Atwood.

                XmlNode result = doc.DocumentElement.SelectSingleNode("/report/activity");
                string externalTestStatus = result.Attributes["result"].Value;

                string externalTestSummary = "";
                XmlNodeList errorMessages = doc.DocumentElement.SelectNodes(".//errmsg");
                if (errorMessages.Count > 0)
                {
                    externalTestSummary = "";
                    foreach (XmlNode error in errorMessages)
                    {
                        externalTestSummary = externalTestSummary + error.InnerText + "\n";
                    }
                }
                else
                {
                    externalTestSummary = externalTestStatus;
                }
                
                //Populate the Test Run object with the results
                if (String.IsNullOrEmpty(automatedTestRun.RunnerName))
                {
                    automatedTestRun.RunnerName = this.ExtensionName;
                }
                automatedTestRun.RunnerTestName = Path.GetFileNameWithoutExtension(runnerTestName);

                //Convert the status for use in SpiraTest
                AutomatedTestRun4.TestStatusEnum executionStatus = AutomatedTestRun4.TestStatusEnum.Passed;
                switch (externalTestStatus)
                {
                    case "Success":
                        executionStatus = AutomatedTestRun4.TestStatusEnum.Passed;
                        break;
                    case "Failed":
                    case "Error":
                        executionStatus = AutomatedTestRun4.TestStatusEnum.Failed;
                        break;
                    case "Warn":
                        executionStatus = AutomatedTestRun4.TestStatusEnum.Caution;
                        break;
                    default:
                        executionStatus = AutomatedTestRun4.TestStatusEnum.Blocked;
                        break;
                }

                //Specify the start/end dates
                automatedTestRun.StartDate = startDate;
                automatedTestRun.EndDate = endDate;

                //The result log
                automatedTestRun.ExecutionStatus = executionStatus;
                automatedTestRun.RunnerMessage = externalTestSummary;
                automatedTestRun.RunnerStackTrace = externalTestDetailedResults;
                automatedTestRun.Format = AutomatedTestRun4.TestRunFormat.PlainText;

                //Now get the detailed activity as 'test steps'
                //Also override the status if we find any failures or warnings
                XmlNodeList resultItems = doc.SelectNodes("/report/activity[@type='root']//activity/item");
                if (resultItems != null && resultItems.Count > 0)
                {
                    automatedTestRun.TestRunSteps = new List<TestRunStep4>();
                    int position = 1;
                    foreach (XmlNode xmlItem in resultItems)
                    {
                        string category = xmlItem.Attributes["category"].Value;
                        string level = xmlItem.Attributes["level"].Value;
                        string message = "";
                        XmlNode xmlMessage = xmlItem.SelectSingleNode("message");
                        if (xmlMessage != null)
                        {
                            message = xmlMessage.InnerText;
                        }

                        //Create the test step
                        TestRunStep4 testRunStep = new TestRunStep4();
                        testRunStep.Description = category;
                        testRunStep.ExpectedResult = "";
                        testRunStep.ActualResult = message;
                        testRunStep.SampleData = "";

                        //Convert the status to the appropriate enumeration value
                        testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.NotRun;
                        switch (level)
                        {
                            case "Success":
                                testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.Passed;
                                break;

                            case "Info":
                                testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.NotApplicable;
                                break;

                            case "Warn":
                                testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.Caution;
                                if (automatedTestRun.ExecutionStatus == AutomatedTestRun4.TestStatusEnum.Passed ||
                                    automatedTestRun.ExecutionStatus == AutomatedTestRun4.TestStatusEnum.NotRun ||
                                    automatedTestRun.ExecutionStatus == AutomatedTestRun4.TestStatusEnum.NotApplicable)
                                {
                                    automatedTestRun.ExecutionStatus = AutomatedTestRun4.TestStatusEnum.Caution;
                                }
                                break;

                            case "Failure":
                            case "Error":
                                testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.Failed;
                                if (automatedTestRun.ExecutionStatus != AutomatedTestRun4.TestStatusEnum.Failed)
                                {
                                    automatedTestRun.ExecutionStatus = AutomatedTestRun4.TestStatusEnum.Failed;
                                }
                                break;
                        }

                        //Add the test step
                        testRunStep.Position = position++;
                        automatedTestRun.TestRunSteps.Add(testRunStep);
                    }
                }
                
                //Report as complete               
                base.status = EngineStatus.OK;
                return automatedTestRun;
            }
            catch (Exception exception)
            {
                //Log the error and denote failure
                LogEvent(exception.Message + " (" + exception.StackTrace + ")", EventLogEntryType.Error);

                //Report as completed with error
                base.status = EngineStatus.Error;
                throw exception;
            }
        }

        private void CreateDirectory(DirectoryInfo directory) {
            if (!directory.Parent.Exists)
            {
                CreateDirectory(directory.Parent);
            } 
            directory.Create();
        }

        private  string GetTimestamp(DateTime value) { return value.ToString("yyyyMMdd_HHmmss"); } 
    }
}
