using System;
using System.Configuration;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using Microsoft.InformationProtection;

namespace RmsFileWatcher
{
    public partial class FormRmsFileWatcher : Form
    {
        private const int           waitPeriodToProcess = 5000;        // 5s between notification and processing, minimum

        FileWatchEngine             fileWatchEngine;
        Action                      action;
        IEnumerable<Microsoft.InformationProtection.Label> labels;
        Microsoft.InformationProtection.Label currentProtectionPolicy;

        // configuration file parameters

        private const string        settingPolicy = "Policy";
        private const string        settingPathCount = "PathCount";
        private const string        settingUnprotectPathCount = "UnprotectPathCount";
        private const string        settingPath = "Path";
        private const string        settingUnprotectPath = "UnprotectPath";

        public FormRmsFileWatcher(ApplicationInfo appInfo)
        {
            InitializeComponent();

            buttonCollapseLog.Tag = false;
            labels = null;
            currentProtectionPolicy = null;

            fileWatchEngine = new FileWatchEngine();
            fileWatchEngine.MillisecondsBeforeProcessing = waitPeriodToProcess;
            fileWatchEngine.EngineEvent += fileWatchEngine_EngineEvent;

            // Initialize Action class, passing in AppInfo.
            action = new Action(appInfo);

            populatePolicyList();

            setFormStateFromConfiguration();
        }

        private void buttonCollapseLog_Click(object sender, EventArgs e)
        {
            int     collapsibleDimension;

            collapsibleDimension = (this.textBoxLog.Location.Y + this.textBoxLog.Height) - (this.buttonCollapseLog.Location.Y + this.buttonCollapseLog.Height);

            // form is currently collapsed, expand it by the size of the controls that are hidden

            if ((bool)buttonCollapseLog.Tag)
            {
                this.MinimumSize = new System.Drawing.Size(372, 430);
                this.Height += collapsibleDimension;
                this.buttonCollapseLog.Image = global::RmsFileWatcher.Properties.Resources.Collapse;
            }

            // form is currently expanded, collapse it by the size of the controls that will be hidden

            else
            {
                this.MinimumSize = new System.Drawing.Size(372, 205);
                this.Height -= collapsibleDimension;
                this.buttonCollapseLog.Image = global::RmsFileWatcher.Properties.Resources.Expand;
            }

            buttonCollapseLog.Tag = !(bool)buttonCollapseLog.Tag;
        }

        /// <summary>
        /// Handle Add folder button, select a folder and add it to the watched folders
        /// list.
        /// </summary>
        private void buttonAdd_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog     dialogFolderBrowser;
            DialogResult            result;

            dialogFolderBrowser = new FolderBrowserDialog();
            dialogFolderBrowser.Description = "Choose a folder to watch...";
            
            result = dialogFolderBrowser.ShowDialog();
            if (result == DialogResult.OK)
            {
                fileWatchEngine.AddWatchedDirectory(dialogFolderBrowser.SelectedPath);
                listBoxWatch.Items.Add(dialogFolderBrowser.SelectedPath);
            }
        }

        /// <summary>
        /// Handle Delete folder button, remove a selected folder from the watched folders
        /// list.
        /// </summary>
        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (listBoxWatch.SelectedIndex != -1)
            {
                fileWatchEngine.RemoveWatchedDirectory((string)listBoxWatch.SelectedItem);
                listBoxWatch.Items.Remove(listBoxWatch.SelectedItem);
            }
        }

        /// <summary>
        /// Handle Delete folder button, remove a selected folder from the watched folders
        /// list.
        /// </summary>
        private void buttonDeleteUnprotect_Click(object sender, EventArgs e)
        {
            if (listBoxUnprotect.SelectedIndex != -1)
            {
                fileWatchEngine.RemoveWatchedDirectory((string)listBoxUnprotect.SelectedItem);
                listBoxUnprotect.Items.Remove(listBoxUnprotect.SelectedItem);
            }
        }

        /// <summary>
        /// Handle Play button to start processing changes in the watched folders list.
        /// </summary>
        private void buttonPlayPause_Click(object sender, EventArgs e)
        {
            // currently not watching for changes

            if (fileWatchEngine.WatchState == WatchState.Suspended)
            {
                this.buttonPlayPause.Image = global::RmsFileWatcher.Properties.Resources.Pause;
                fileWatchEngine.StartWatching();
                timerProcessChanges.Enabled = true;
            }

            // currently watching for changes

            else
            {
                this.buttonPlayPause.Image = global::RmsFileWatcher.Properties.Resources.Play;
                fileWatchEngine.SuspendWatching();
                timerProcessChanges.Enabled = false;
            }
        }

        private void buttonUndo_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialogFolderBrowser;
            DialogResult result;

            dialogFolderBrowser = new FolderBrowserDialog();
            dialogFolderBrowser.Description = "Choose a folder to watch for files to un-protect...";

            result = dialogFolderBrowser.ShowDialog();
            if (result == DialogResult.OK)
            {
                fileWatchEngine.AddWatchedDirectory(dialogFolderBrowser.SelectedPath);
                listBoxUnprotect.Items.Add(dialogFolderBrowser.SelectedPath);
            }
        }

        /// <summary>
        /// Handle Form Closing event to save configuration state.
        /// </summary>
        private void FormFileWatcher_FormClosing(object sender, FormClosingEventArgs e)
        {
            string[]    pathsToWatch;
            string[]    pathsToUnprotect;
            string      policyToApply;

            readConfigurationFromFormState(out pathsToWatch, out policyToApply, out pathsToUnprotect);
            saveConfiguration(pathsToWatch, policyToApply, pathsToUnprotect);
        }

        /// <summary>
        /// Handle timer tick event to processed accumulated changes in watched folders
        /// list.
        /// </summary>
        private void timerProcessChanges_Tick(object sender, EventArgs e)
        {
            // only process changes that happened more than x seconds ago to try
            // to handle boundary cases where a change triggers multiple notifications
            // and the timer goes off in between the notifications

            if (comboBoxTemplates.SelectedIndex <= 0)
            {
                this.Invoke(new AppendToLog(doAppendToLog), "Can't protect files until a protection policy is specified\r\n");
                return;
            }

            try
            {
                timerProcessChanges.Enabled = false;
                fileWatchEngine.ProcessWatchedChanges();
            }
            finally
            {
                timerProcessChanges.Enabled = true;
            }
        }

        /// <summary>
        /// Handle File Watch Engine events for state changes and failures.
        /// </summary>
        private void fileWatchEngine_EngineEvent(object sender, EngineEventArgs e)
        {
            if (e.NotificationType == EngineNotificationType.Watching ||
                e.NotificationType == EngineNotificationType.Suspended)
            {
                this.Invoke(new AppendToLog(doAppendToLog), "** " + e.NotificationType.ToString() + "\r\n");
            }
            else if (e.NotificationType == EngineNotificationType.Processing)
            {
                this.Invoke(new AppendToLog(doAppendToLog), e.NotificationType.ToString() + ": " + e.FullPath + "...");

                // append 'p' to file extension to denote as protected
                //String[] outputPath = e.FullPath.Split('.');
                //outputPath[outputPath.Length - 1] = "p" + outputPath[outputPath.Length - 1];
                //String outFile = String.Join(".", outputPath);
                String[] outputPath = e.FullPath.Split('\\');

                // create 'protected' dir if it doesn't exist
                //Directory.CreateDirectory();

                //outputPath[outputPath.Length - 1] = "protected\\" + outputPath[outputPath.Length - 1];
                //String outFile = String.Join("\\", outputPath);

                String tmpFilePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

                //String outFile = e.FullPath;
                Action.FileOptions options = new Action.FileOptions
                {
                    FileName = e.FullPath,
                    OutputName = tmpFilePath,
                    ActionSource = ActionSource.Manual,
                    AssignmentMethod = AssignmentMethod.Standard,
                    DataState = DataState.Rest,
                    GenerateChangeAuditEvent = true,
                    IsAuditDiscoveryEnabled = true,
                    LabelId = currentProtectionPolicy.Id
                };

                this.Invoke(new AppendToLog(doAppendToLog), "Checking for existing Label" + "\r\n");

                ContentLabel currLabel = action.GetLabel(options);

                Boolean doProtect = true;
                // determine if this file should be protected or unprotected
                if(listBoxUnprotect.Items.Count > 0)
                { 
                    for (int i = 0; i < listBoxUnprotect.Items.Count; i++)
                    {
                        String path = (string)listBoxUnprotect.Items[i];
                        if(e.FullPath.StartsWith(path))
                        {
                            doProtect = false;
                            break;
                        }
                    }
                }

                ////
                if (doProtect)
                {
                    if (currLabel != null)
                        this.Invoke(new AppendToLog(doAppendToLog), "Already Protected!\r\n");

                    if (currentProtectionPolicy != null && currLabel == null)
                    {
                        this.Invoke(new AppendToLog(doAppendToLog), "Setting Label and saving to " + tmpFilePath + "\r\n");
                        action.SetLabel(options);
                        this.Invoke(new AppendToLog(doAppendToLog), "Replacing file with protected version" + "\r\n");
                        File.Delete(e.FullPath);
                        this.Invoke(new AppendToLog(doAppendToLog), "Deleted old " + "\r\n");
                        File.Copy(tmpFilePath, e.FullPath);
                        this.Invoke(new AppendToLog(doAppendToLog), "Protected!\r\n");
                    }

                } else { // Unprotect the file

                    if (currLabel == null)
                        this.Invoke(new AppendToLog(doAppendToLog), "Already NOT Protected!\r\n");

                    this.Invoke(new AppendToLog(doAppendToLog), "Unprotect this file \r\n");
                    options.LabelId = null; // remove label??
                    this.Invoke(new AppendToLog(doAppendToLog), "Removing Label and saving to " + tmpFilePath + "\r\n");
                    action.DeleteLabel(options);
                    this.Invoke(new AppendToLog(doAppendToLog), "Replacing file with unprotected version" + "\r\n");
                    File.Delete(e.FullPath);
                    this.Invoke(new AppendToLog(doAppendToLog), "Deleted old " + "\r\n");
                    File.Copy(tmpFilePath, e.FullPath);
                    this.Invoke(new AppendToLog(doAppendToLog), "Protection Removed!\r\n");
                }
            }
            else
            {
                this.Invoke(new AppendToLog(doAppendToLog), e.NotificationType.ToString() + "\r\n");
            }
        }

        /// <summary>
        /// Handle protection policy selection change to track policy to apply to files.
        /// </summary>
        private void comboBoxTemplates_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxTemplates.SelectedIndex != -1)
            {
                currentProtectionPolicy = findLabel((string)comboBoxTemplates.SelectedItem);
            }
        }

        //
        // Private helper functions
        //

        /// <summary>
        /// Load configuration and populate the form's UI from this configuration.
        /// </summary>
        private void setFormStateFromConfiguration()
        {
            string[]    pathsToWatch;
            string[]    pathsToUnprotect;
            string      policyToApply;

            loadConfiguration(out pathsToWatch, out policyToApply, out pathsToUnprotect);

            if (pathsToWatch != null)
            {
                foreach (string s in pathsToWatch)
                {
                    listBoxWatch.Items.Add(s);
                    fileWatchEngine.AddWatchedDirectory(s);
                }
            }
            if (pathsToUnprotect != null)
            {
                foreach (string s in pathsToUnprotect)
                {
                    listBoxUnprotect.Items.Add(s);
                    fileWatchEngine.AddWatchedDirectory(s);
                }
            }
            if (policyToApply != "")
            {
                comboBoxTemplates.SelectedItem = policyToApply;
            }
        }

        /// <summary>
        /// Read configuration from the current form state.
        /// </summary>
        private void readConfigurationFromFormState(out string[] pathsToWatch, out string policyToApply, out string[] pathsToUnprotect)
        {
            pathsToWatch = null;
            pathsToUnprotect = null;
            if (listBoxWatch.Items.Count > 0)
            {
                pathsToWatch = new string[listBoxWatch.Items.Count];
                for (int i = 0; i < listBoxWatch.Items.Count; i++)
                {
                    pathsToWatch[i] = (string)listBoxWatch.Items[i];
                }
            }

            if (listBoxUnprotect.Items.Count > 0)
            {
                pathsToUnprotect = new string[listBoxUnprotect.Items.Count];
                for (int i = 0; i < listBoxUnprotect.Items.Count; i++)
                {
                    pathsToUnprotect[i] = (string)listBoxUnprotect.Items[i];
                }
            }
            policyToApply = "";
            if (comboBoxTemplates.SelectedIndex > 0)
            {
                policyToApply = (string)comboBoxTemplates.SelectedItem;
            }
        }

        /// <summary>
        /// Query for available protection policies and fill the policy combo box for
        /// selection.
        /// </summary>
        private void populatePolicyList()
        {

            // List all labels available to the engine created in Action
            labels = action.ListLabels();
            

            comboBoxTemplates.BeginUpdate();
            comboBoxTemplates.Items.Add("-- Choose a policy --");

            foreach (Microsoft.InformationProtection.Label label in labels)
            {
                comboBoxTemplates.Items.Add(label.Name);
            }

            comboBoxTemplates.SelectedIndex = 0;
            comboBoxTemplates.EndUpdate();
        }

        private Microsoft.InformationProtection.Label findLabel(string s)
        {
            Microsoft.InformationProtection.Label item;

            item = null;
            foreach (Microsoft.InformationProtection.Label label in labels)
            {
                if (label.Name == s)
                {
                    item = label;
                }
            }

            return item;
        }

        private delegate void AppendToLog(string text);

        private void doAppendToLog(string text)
        {
            this.textBoxLog.AppendText(text);
        }

        /// <summary>
        /// Load state from application configuration file.
        /// </summary>
        private void loadConfiguration(out string[] pathsToWatch, out string policyToApply, out string[] pathsToUnprotect)
        {
            NameValueCollection nvc;

            policyToApply = "";
            pathsToWatch = null;
            pathsToUnprotect = null;

            nvc = (NameValueCollection)ConfigurationManager.AppSettings;
            if (nvc.AllKeys.Contains(settingPolicy))
            {
                policyToApply = nvc[settingPolicy];
            }

            if (nvc.AllKeys.Contains(settingPathCount))
            {
                int pathCount;

                pathCount = System.Convert.ToInt32(nvc[settingPathCount]);
                pathsToWatch = new string[pathCount];
                for (int i = 0; i < pathCount; i++)
                {
                    string key;

                    key = settingPath + i.ToString();
                    if (nvc.AllKeys.Contains(key))
                    {
                        pathsToWatch[i] = nvc[key];
                    }
                }
            }
            if (nvc.AllKeys.Contains(settingUnprotectPathCount))
            {
                int pathCount;

                pathCount = System.Convert.ToInt32(nvc[settingUnprotectPathCount]);
                pathsToUnprotect = new string[pathCount];
                for (int i = 0; i < pathCount; i++)
                {
                    string key;

                    key = settingUnprotectPath + i.ToString();
                    if (nvc.AllKeys.Contains(key))
                    {
                        pathsToUnprotect[i] = nvc[key];
                    }
                }
            }
        }

        /// <summary>
        /// Save state to the application configuration file.
        /// </summary>
        private void saveConfiguration(string[] pathsToWatch, string policyToApply, string[] pathsToUnprotect)
        {
            Configuration       appConfig;
            AppSettingsSection  appSettings;

            appConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            appSettings = appConfig.AppSettings;
            NameValueCollection nvc = (NameValueCollection)ConfigurationManager.AppSettings;
            int prevPathCount = System.Convert.ToInt32(nvc[settingPathCount]);
            for (int i = 0; i < prevPathCount; i++)
            {
                appSettings.Settings.Remove(settingPath + i.ToString());
            }
            //Unprotect
            prevPathCount = System.Convert.ToInt32(nvc[settingUnprotectPathCount]);
            for (int i = 0; i < prevPathCount; i++)
            {
                appSettings.Settings.Remove(settingUnprotectPath + i.ToString());
            }
            appSettings.Settings.Remove(settingPolicy);
            appSettings.Settings.Remove(settingPathCount);
            appSettings.Settings.Remove(settingUnprotectPathCount);

            appSettings.Settings.Add(settingPolicy, policyToApply);
            //protect these
            if (pathsToWatch != null)
            {
                appSettings.Settings.Add(settingPathCount, pathsToWatch.Length.ToString());
                for (int i = 0; i < pathsToWatch.Length; i++)
                {
                    appSettings.Settings.Add(settingPath + i.ToString(), pathsToWatch[i]);
                }
            }
            //unprotect these
            if (pathsToUnprotect != null)
            {
                appSettings.Settings.Add(settingUnprotectPathCount, pathsToUnprotect.Length.ToString());
                for (int i = 0; i < pathsToUnprotect.Length; i++)
                {
                    appSettings.Settings.Add(settingUnprotectPath + i.ToString(), pathsToUnprotect[i]);
                }
            }

            appConfig.Save(ConfigurationSaveMode.Modified);
        }
    }
}
