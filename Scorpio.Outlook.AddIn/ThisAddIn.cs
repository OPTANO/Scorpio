#region Copyright (c) ORCONOMY GmbH 

// ////////////////////////////////////////////////////////////////////////////////
//                                                                   
//        ORCONOMY GmbH Source Code                                   
//        Copyright (c) 2010-2016 ORCONOMY GmbH                       
//        ALL RIGHTS RESERVED.                                        
//                                                                    
//    The entire contents of this file is protected by German and       
//    International Copyright Laws. Unauthorized reproduction,        
//    reverse-engineering, and distribution of all or any portion of  
//    the code contained in this file is strictly prohibited and may  
//    result in severe civil and criminal penalties and will be       
//    prosecuted to the maximum extent possible under the law.        
//                                                                    
//    RESTRICTIONS                                                    
//                                                                    
//    THIS SOURCE CODE AND ALL RESULTING INTERMEDIATE FILES           
//    ARE CONFIDENTIAL AND PROPRIETARY TRADE SECRETS OF               
//    ORCONOMY GMBH. 
//                                                                    
//    THE SOURCE CODE CONTAINED WITHIN THIS FILE AND ALL RELATED      
//    FILES OR ANY PORTION OF ITS CONTENTS SHALL AT NO TIME BE        
//    COPIED, TRANSFERRED, SOLD, DISTRIBUTED, OR OTHERWISE MADE       
//    AVAILABLE TO OTHER INDIVIDUALS WITHOUT WRITTEN CONSENT  
//    AND PERMISSION FROM ORCONOMY GMBH.                              
//                                                                   
// ////////////////////////////////////////////////////////////////////////////////

#endregion

using Office = Microsoft.Office.Core;

namespace Scorpio.Outlook.AddIn
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    using System.Threading.Tasks;
    using System.Windows.Interop;

    using log4net;

    using Microsoft.Office.Interop.Outlook;
    using Microsoft.Office.Tools;

    using Scorpio.Outlook.AddIn.Extensions;
    using Scorpio.Outlook.AddIn.Helper;
    using Scorpio.Outlook.AddIn.Misc;
    using Scorpio.Outlook.AddIn.Report;
    using Scorpio.Outlook.AddIn.Synchronization;
    using Scorpio.Outlook.AddIn.UserInterface.Controls;
    using Scorpio.Outlook.AddIn.UserInterface.RibbonBars;
    using Scorpio.Outlook.AddIn.UserInterface.ViewModel;

    using Exception = System.Exception;

    /// <summary>
    /// The Scorpio Outlook plugin.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>
        /// The logger.
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(ThisAddIn));

        /// <summary>
        /// Name of the calendar which holds the time entries from redmine.
        /// </summary>
#if DEBUG
        public static readonly string CalendarName = "Redmine-Debug";
#else
            public static readonly string CalendarName = "Redmine";
        #endif

        #region Private fields

        /// <summary>
        /// The ribbon bar for this plugin.
        /// </summary>
        private ScorpioRibbon _ribbon;

        /// <summary>
        /// The report creator which creates the monthly report.
        /// </summary>
        private ReportCreator _reportCreator;

        /// <summary>
        /// The custom task pane as known by outlook.
        /// </summary>
        private CustomTaskPane _customTaskPane;

        #endregion

        #region Properties

        /// <summary>
        /// Gets the calendar in outlook which is the target for redmine time entries
        /// </summary>
        public MAPIFolder RedmineCalendar { get; private set; }

        /// <summary>
        /// Gets the redmine synchronizer instance
        /// </summary>
        public Synchronizer Synchronizer { get; private set; }

        /// <summary>
        /// Gets the sync state object which keeps information about the synchronization process.
        /// </summary>
        public SyncState SyncState { get; private set; }

        /// <summary>
        /// Gets the calendar state object. This keeps information about the currently opened calendar.
        /// </summary>
        public CalendarState CalendarState { get; private set; }

        /// <summary>
        /// Gets or sets the viewmodel for the scorpio task pane.
        /// </summary>
        public ScorpioTaskPaneViewModel ScorpioViewModel { get; set; }

        /// <summary>
        /// Gets the ui synchronizer which is updating ui infos displayed to the users
        /// </summary>
        public UiUserInfoSynchronizer UiUserInfoSynchronizer { get; private set; }

        #endregion

        #region Public addin functions

        /// <summary>
        /// Opens the calendar with redmine time entries
        /// </summary>
        public void OpenCalendar()
        {
            var primaryCalendar = this.Application.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            Application.ActiveExplorer().SelectFolder(primaryCalendar.Folders[CalendarName]);
            Application.ActiveExplorer().CurrentFolder.Display();

            var view = primaryCalendar.CurrentView;
        }

        /// <summary>
        /// Reconnects to the redmine system
        /// </summary>
        public void ReconnectToRedmine()
        {
            this.Synchronizer.InitializeRedmineNew();
        }

        #endregion

        /// <summary>
        /// Triggers the calculations for the monthly report.
        /// </summary>
        public void CalculateMonthlyReport()
        {
            this._reportCreator.CalculateMonthlyReport(this.Synchronizer, this.Application);
        }

        /// <summary>
        /// Shows the scorpio task pane.
        /// </summary>
        public void ShowTaskPane()
        {
            this._customTaskPane.Visible = true;
        }

        /// <summary>
        /// Implements a check if a request to redmine shall be allowed. This method allows any requests 
        /// to the test redmine URL regardless of certificate validation errors.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="cert">The certificate</param>
        /// <param name="chain">The certificate validation chain</param>
        /// <param name="policyErrors">The policy errors</param>
        /// <returns>true if the request url belongs to the test redmine, false otherwise.</returns>
        private static bool AllowTestRedmineCertificate(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors policyErrors)
        {
            var request = sender as HttpWebRequest;
            if (request != null && request.Address.AbsoluteUri.ToUpper().Contains("192.168.0.93"))
            {
                return true;
            }

            return policyErrors == SslPolicyErrors.None;
        }

        /// <summary>
        /// Method that is called when the addin is started.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The event arguments</param>
        private void ThisAddInStartup(object sender, EventArgs e)
        {
            try
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                Log.Info("SCORPIO plugin starting initialization");
#if DEBUG

                // If we are in debug, allow the incorrect certificate that the test-redmine at 192.168.0.93 has.
                ServicePointManager.ServerCertificateValidationCallback += AllowTestRedmineCertificate;
#endif

                // Configure Log4net
                log4net.Config.XmlConfigurator.Configure();
                Log.Info(string.Format("SCORPIO plugin configured Log4Net after {0} ms", stopwatch.ElapsedMilliseconds));

                this._reportCreator = new ReportCreator();

                // init the plugin logic
                this.CheckRequirements(this.Application.ActiveExplorer());
                Log.Info(string.Format("SCORPIO plugin created necessary Outlook Objects after {0} ms", stopwatch.ElapsedMilliseconds));
                
                // Initialization of the task pane takes about 2 seconds. Therefore, schedule it to the ui 
                // thread for later execution to not slow down plugin startup times.
                var uiScheduler = TaskScheduler.FromCurrentSynchronizationContext();
                var taskPaneTask = new Task(this.CreateCustomTaskPane);
                taskPaneTask.Start(uiScheduler);

                Log.Info(string.Format("SCORPIO plugin created task pane after {0} ms", stopwatch.ElapsedMilliseconds));
                Log.Info(string.Format("SCORPIO plugin successfully initialized in {0} ms", stopwatch.ElapsedMilliseconds));

                this.ReconnectToRedmine();
                Log.Info(string.Format("SCORPIO plugin triggered connection to redmine after {0} ms", stopwatch.ElapsedMilliseconds));

                stopwatch.Stop();
            }
            catch (Exception ex)
            {
                Log.Error("SCORPIO plugin could not be initialized.", ex);
                throw;
            }
        }

        /// <summary>
        /// Method that creates and initializes the custom task pane.
        /// </summary>
        private void CreateCustomTaskPane()
        {
            try
            {
                var taskPaneContainer = new ScorpioTaskPaneContainer();
                this._customTaskPane = this.CustomTaskPanes.Add(taskPaneContainer, "SCORPIO");
                this._customTaskPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                this._customTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                this._customTaskPane.Visible = true;
                this._customTaskPane.Width = 118;
                this.ScorpioViewModel = new ScorpioTaskPaneViewModel();
                taskPaneContainer.TaskPane.DataContext = this.ScorpioViewModel;
            }
            catch (Exception ex)
            {
                Log.Error("SCORPIO could not initialize the task pane.", ex);
                throw;
            }
        }

        /// <summary>
        /// Shows the settings dialog to the user.
        /// </summary>
        internal void ShowSettingsDialog()
        {
            try
            {
                // create dialog and show it
                var dialog = new SettingsEditorDialog();
                dialog.ShowDialog();

                // handle positive result from closing the dialog
                if (dialog.DialogResult.HasValue && dialog.DialogResult.Value)
                {
                    Globals.ThisAddIn.ReconnectToRedmine();
                    this.UiUserInfoSynchronizer.RestartTimer(dialog.RefreshTime);
                }
            }
            catch (Exception ex)
            {
                Log.Error("Error while opening the settings dialog: ", ex);
                throw;
            }
        }

        /// <summary>
        /// Called when Outlook shuts down.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The event arguments</param>
        private void ThisAddInShutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            // must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        /// <summary>
        /// Method that opens the redmine calendar, creates a new time entry in the calendar, and opens the time entry for editing.
        /// </summary>
        internal void CreateNewTimeEntry()
        {
            this.OpenCalendar();

            try
            {
                var newAppointment = (AppointmentItem)this.RedmineCalendar.Items.Add(OlItemType.olAppointmentItem);
                newAppointment.Start = DateTime.Now.Date.AddHours(DateTime.Now.Hour - 1);
                newAppointment.End = DateTime.Now.Date.AddHours(DateTime.Now.Hour);
                newAppointment.AllDayEvent = false;
                newAppointment.ReminderSet = false;

                newAppointment.Save();
                newAppointment.Display(true);
            }
            catch (Exception ex)
            {
                Log.Error("Could not create new appointment in redmine calendar. ", ex);
            }
        }

        /// <summary>
        /// Creates the ribbon object for this addin
        /// </summary>
        /// <returns>The ribbon object</returns>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return this._ribbon ?? (this._ribbon = new ScorpioRibbon());
        }

        /// <summary>
        /// Creates recurring time entries.
        /// </summary>
        internal void CreateRecurringTimeEntries()
        {
            this.Synchronizer.CreateRecurringTimeEntries();
        }

        /// <summary>
        /// Method that initializes the Outlook plugin, by checking if all necessary elements (s.a. the redmine calendar) exist. 
        /// It creates those necessary elements if they do not already exist.
        /// </summary>
        /// <param name="currentExplorer">The current explorer</param>
        private void CheckRequirements(Explorer currentExplorer)
        {
            // create the calendar folder
            var primaryCalendar = currentExplorer.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            this.RedmineCalendar = OutlookHelper.CreateOrGetFolder(primaryCalendar, CalendarName, OlDefaultFolders.olFolderCalendar);
            OutlookHelper.CreateScorpioUserDefinedProperties(this.RedmineCalendar);
            OutlookHelper.CreateScorpioCategories();

            this._reportCreator.CheckRequirements(currentExplorer);

            // create new state objects
            this.CalendarState = new CalendarState();
            this.SyncState = new SyncState();

            // create new sync objects
            this.Synchronizer = new Synchronizer(this.RedmineCalendar);
            Func<DateTime, DateTime, List<AppointmentItem>> getAppointmentsFunction =
                (start, end) => this.Synchronizer.Calendar.GetAppointmentsInRange(start, end, includeEnd: false);
            this.UiUserInfoSynchronizer = new UiUserInfoSynchronizer(getAppointmentsFunction);

            // add listener
            // ui update listener
            this.Synchronizer.AppointmentChanged += (sender, args) => this.UiUserInfoSynchronizer.HandleAppointmentChange(sender, args);

            // sync status listener
            this.SyncState.ConnectionStateChanged += (sender, args) =>
                {
                    if (this._ribbon != null)
                    {
                        this._ribbon.UpdateConnectionStatus();
                    }
                };
            this.SyncState.StatusChanged += (sender, args) =>
                {
                    if (this._ribbon != null)
                    {
                        this._ribbon.UpdateStatus();
                    }
                };
        }
        
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += this.ThisAddInStartup;
            this.Shutdown += this.ThisAddInShutdown;
        }

        #endregion
    }
}