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

namespace Scorpio.Outlook.AddIn.UserInterface.RibbonBars
{
    using System;
    using System.Deployment.Application;
    using System.Drawing;
    using System.Threading.Tasks;
    using System.Windows;

    using Microsoft.Office.Core;

    using Scorpio.Outlook.AddIn.UserInterface.View;
    using Scorpio.Outlook.AddIn.UserInterface.ViewModel;

    /// <summary>
    /// Partial class for the scorpio ribbon
    /// </summary>
    public partial class ScorpioRibbon
    {
        #region Public Methods and Operators

        /// <summary>
        /// Gets the enabled state for controls that trigger synchronization with redmine.
        /// </summary>
        /// <param name="control">The control for which to determine if it should be enabled.</param>
        /// <returns><code>true</code> if the control should be enabled, <code>false</code> otherwise.</returns>
        public bool GetConnectEnabled(IRibbonControl control)
        {
            if (control.Id == "connectRedmine")
            {
                return !Globals.ThisAddIn.Synchronizer.IsConnecting;
            }
            if (control.Id == "resetTimeEntries" || control.Id == "saveTimeEntries")
            {
                return Globals.ThisAddIn.Synchronizer.CanSyncTimeEntries && Globals.ThisAddIn.CalendarState.CalendarView != null;
            }
            if (control.Id == "createreport")
            {
                // TODO DS: reenable later. Used to be the same as for resetTimeEntries and saveTimeEntries.
                return false;
            }
            return false;
        }

        /// <summary>
        /// Gets the label-string for the connection label.
        /// </summary>
        /// <param name="control">The control for which to get the label</param>
        /// <returns>The label-string.</returns>
        public string GetConnectLabelText(IRibbonControl control)
        {
            if (control.Id == "connectRedmine")
            {
                return "Synchronisiere Redmine";
            }
            return "";
        }

        /// <summary>
        /// Gets an image for a ribbon control.
        /// </summary>
        /// <param name="control">The control for which to get the image.</param>
        /// <returns>The image for the control.</returns>
        public Bitmap GetRibbonImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "saveTimeEntries":
                    return Properties.Resources.diskette;

                case "resetTimeEntries":
                    return Properties.Resources.arrow_undo;

                case "showCalendar":
                    return Properties.Resources.calendar;

                case "connectRedmine":
                    return Properties.Resources.arrow_refresh;

                case "showSettings":
                    return Properties.Resources.setting_tools;

                case "createreport":
                    return Properties.Resources.report_user;
                case "showTaskPane":
                    return Properties.Resources.application_side_expand;
                case "createSingle":
                    return Properties.Resources.date_add;
                case "createRecurring":
                    return Properties.Resources.date_relation;
                case "showHours":
                    return Properties.Resources.report_user;
            }
            return null;
        }

        /// <summary>
        /// Gets the label-string for status and hour labels
        /// </summary>
        /// <param name="control">The control for which to get the label-string</param>
        /// <returns>The label-string for the control</returns>
        public string GetStatusLabel(IRibbonControl control)
        {
            if (control.Id == "statusLabel")
            {
                return Globals.ThisAddIn.SyncState.Status;
            }
            if (control.Id == "hoursLabel")
            {
                return string.Format("Stunden: {0:N2}", Globals.ThisAddIn.SyncState.HoursInView);
            }
            return "";
        }

        /// <summary>
        /// Method that gets the label text for the version label.
        /// </summary>
        /// <param name="control">The ribbon control.</param>
        /// <returns>The deployment version number of the plugin.</returns>
        public string GetVersionLabel(IRibbonControl control)
        {
            Version version;

            try
            {
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
            catch (Exception)
            {
                // In development, there is no Applicationdeployment.
                return "Entwicklung";
            }

            if (version == null)
            {
                return "unbekannt";
            }

            return version.Major + "." + version.Minor + "." + version.Build + "." + version.Revision;
        }

        /// <summary>
        /// Called when the user presses the button for reconnecting to redmine.
        /// </summary>
        /// <param name="control">The control which was pressed.</param>
        public void OnConnect(IRibbonControl control)
        {
            Globals.ThisAddIn.ReconnectToRedmine();
        }

        /// <summary>
        /// Called when the user presses the button for creating recurring time entries.
        /// </summary>
        /// <param name="control">The control which was pressed.</param>
        public void OnCreateRecurring(IRibbonControl control)
        {
            Globals.ThisAddIn.CreateRecurringTimeEntries();
        }

        /// <summary>
        /// Called when the user presses the button for creating a single time entry.
        /// </summary>
        /// <param name="control">The control which was pressed.</param>
        public void OnCreateSingle(IRibbonControl control)
        {
            Globals.ThisAddIn.CreateNewTimeEntry();
        }

        /// <summary>
        /// Called when the user presses the button for creating the monthly report of work time.
        /// </summary>
        /// <param name="control">The control which was pressed.</param>
        public void OnReport(IRibbonControl control)
        {
            // TODO Datum des aktuell gewählten Tag
            Globals.ThisAddIn.CalculateMonthlyReport();
        }

        /// <summary>
        /// Called when the user presses the button for resetting time entries.
        /// </summary>
        /// <param name="control">The control which was pressed.</param>
        /// <returns>An awaitable task.</returns>
        public async Task OnResetTimeEntries(IRibbonControl control)
        {
            await Globals.ThisAddIn.Synchronizer.RevertTimeEntriesToRedmineState();
        }

        /// <summary>
        /// Called when the user presses the button for saving changes to time entries.
        /// </summary>
        /// <param name="control">The control which was pressed.</param>
        /// <returns>An awaitable task.</returns>
        public async Task OnSaveTimeEntries(IRibbonControl control)
        {
            await Globals.ThisAddIn.Synchronizer.SaveTimeEntriesAsync();
        }

        /// <summary>
        /// Called when the user presses the button for showing the redmine calendar.
        /// </summary>
        /// <param name="control">The control which was pressed.</param>
        public void OnShowCalendar(IRibbonControl control)
        {
            Globals.ThisAddIn.OpenCalendar();
        }

        /// <summary>
        /// Called when the user presses the button for showing the settings dialog.
        /// </summary>
        /// <param name="control">The control which was pressed.</param>
        public void OnShowSettings(IRibbonControl control)
        {
            Globals.ThisAddIn.ShowSettingsDialog();
        }

        /// <summary>
        /// Called when the user presses the button for showing the sho time entries dialog.
        /// </summary>
        /// <param name="control">The control which was pressed.</param>
        public void OnShowHours(IRibbonControl control)
        {
            var dialog = new ShowTimeEntries();
            dialog.DataContext = new ShowTimeEntriesViewModel();
            dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            dialog.ShowDialog();
        }

        /// <summary>
        /// Called when the user presses the button for showing the task pane.
        /// </summary>
        /// <param name="control">The control which was pressed.</param>
        public void OnShowTaskpane(IRibbonControl control)
        {
            Globals.ThisAddIn.ShowTaskPane();
        }

        /// <summary>
        /// Invalidates all controls that depend on the connection status to redmine.
        /// </summary>
        public void UpdateConnectionStatus()
        {
            if (this.ribbon == null)
            {
                return;
            }
            this.ribbon.InvalidateControl("connectRedmine");
            this.ribbon.InvalidateControl("saveTimeEntries");
            this.ribbon.InvalidateControl("resetTimeEntries");
            this.ribbon.InvalidateControl("createreport");
        }

        /// <summary>
        /// Invalidates all controls that depend on the status.
        /// </summary>
        public void UpdateStatus()
        {
            if (this.ribbon == null)
            {
                return;
            }
            this.ribbon.InvalidateControl("statusLabel");
            this.ribbon.InvalidateControl("hoursLabel");
        }

        #endregion
    }
}