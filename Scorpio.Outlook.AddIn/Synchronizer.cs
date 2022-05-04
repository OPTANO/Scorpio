#region Copyright (c) ORCONOMY GmbH 

// ////////////////////////////////////////////////////////////////////////////////
//                                                                   
//        ORCONOMY GmbH Source Code                                   
//        Copyright (c) 2010-2017 ORCONOMY GmbH                       
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

namespace Scorpio.Outlook.AddIn
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using System.Windows.Interop;

    using DevExpress.Mvvm.Native;

    using log4net;

    using Microsoft.Office.Interop.Outlook;

    using Scorpio.Outlook.AddIn.Cache;
    using Scorpio.Outlook.AddIn.Extensions;
    using Scorpio.Outlook.AddIn.Helper;
    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Misc;
    using Scorpio.Outlook.AddIn.Properties;
    using Scorpio.Outlook.AddIn.Synchronization;
    using Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource;
    using Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource.Exceptions;
    using Scorpio.Outlook.AddIn.Synchronization.Helper;
    using Scorpio.Outlook.AddIn.UserInterface.Controls;

    using Exception = System.Exception;

    /// <summary>
    /// Class that is responsible for keeping time entries in sync with the calendar in outlook.
    /// </summary>
    public class Synchronizer
    {
        #region Static Fields

        /// <summary>
        /// The logger.
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(Synchronizer));

        /// <summary>
        ///  A description of the regular expression:
        ///  <para>
        ///  Beginning of line or string [1]: A numbered capture group. [#?] #, zero or one repetitions
        ///  [IssueNumber]: A named capture group. [\d+] Any digit, one or more repetitions Any character, any number of repetitions
        /// </para>
        /// <para>
        /// Regular expression built for C# on: Mi, Sep 23, 2015, 05:28:42 
        /// Using Expresso Version: 3.0.4750, http://www.ultrapico.com
        /// </para>
        /// </summary>
        private static readonly Regex RxIssueNumber = new Regex("^(#?)(?<IssueNumber>\\d+).*", RegexOptions.CultureInvariant | RegexOptions.Compiled);

        #endregion

        #region Fields

        /// <summary>
        /// List of all AppointmentItems in the outlook calendar. Elements in this set have a deletionlistener and a changelistener attached.
        /// </summary>
        private readonly HashSet<AppointmentItem> _managedItems = new HashSet<AppointmentItem>();

        /// <summary>
        /// A dictionary of all time entries that the user can read in the external source. Mapping from time entry id to time entry.
        /// </summary>
        private IDictionary<int, ActivityInfo> _activities = new Dictionary<int, ActivityInfo>();

        /// <summary>
        /// The manager for the external data source
        /// </summary>
        private IExternalSource _externalDataSource;

        /// <summary>
        /// A dictionary of all issues that the user can read in the external source. Mapping from issue id to issue.
        /// </summary>
        private IDictionary<int, IssueInfo> _issues = new Dictionary<int, IssueInfo>();

        /// <summary>
        /// The number of last used issues to display
        /// </summary>
        private int _numberLastUsedIssues;

        /// <summary>
        /// A dictionary of all projects that the user can read in the external source. Mapping from project id to project.
        /// </summary>
        private IDictionary<int, ProjectInfo> _projects = new Dictionary<int, ProjectInfo>();

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Synchronizer"/> class.
        /// </summary>
        /// <param name="calendar">The target calendar in outlook for syncing time entries</param>
        public Synchronizer(MAPIFolder calendar)
        {
            this.Calendar = calendar;
            Globals.ThisAddIn.SyncState.Status = "Nicht verbunden";
            this.LastUsedIssues = new List<IssueInfo>();
            this.FavoriteIssues = new List<IssueInfo>();

            // At startup, register all appointments in the calendar for change tracking.
            foreach (var item in this.Calendar.Items)
            {
                var appointment = item as AppointmentItem;
                if (appointment != null)
                {
                    this.RegisterAppointment(appointment);
                }
            }

            this.CalendarItems = this.Calendar.Items;

            this.CalendarItems.ItemAdd += this.OnItemsOnItemAdd;
            this._numberLastUsedIssues = Math.Max(Settings.Default.NumberLastUsedIssues, 0);
        }

        #endregion

        #region Public Events

        /// <summary>
        /// Event is raised when an appointment has changed
        /// </summary>
        public event EventHandler AppointmentChanged;

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Build the issue information string (link + name)
        /// </summary>
        /// <param name="issueId">The issue id</param>
        /// <returns>The link and the name of the issue</returns>
        public Tuple<string, string> BuildIssueInformation(int issueId)
        {
            var issueLink = string.Format("{0}/issues/{1}", this.ConnectionUrl, issueId);
            var issueName = string.Format("Issue {0}", issueId);

            if (this._issues.ContainsKey(issueId))
            {
                issueName = this._issues[issueId].Name;
            }
            return new Tuple<string, string>(issueLink, issueName);
        }

        /// <summary>
        /// Build the project information (link + name)
        /// </summary>
        /// <param name="projectId">The project id</param>
        /// <returns>The link and the name of the project</returns>
        public Tuple<string, string> BuildProjectInformation(int projectId)
        {
            var projectLink = string.Format("{0}/projects/{1}", this.ConnectionUrl, projectId);
            var projectName = string.Format("Project {0}", projectId);

            if (this._projects.ContainsKey(projectId))
            {
                projectName = this._projects[projectId].Name;
            }
            return new Tuple<string, string>(projectLink, projectName);
        }

        /// <summary>
        /// Method to get all issues assigned to me
        /// </summary>
        /// <returns>the list of all issues assigned to me</returns>
        public List<IssueInfo> GetAllIssuesAssignedToMe()
        {
            var parameters = new DataSourceParameter() { AssignedToUserId = -1, };
            var issuesForMe = this._externalDataSource.GetIssueInfoList(parameters);
            return issuesForMe.ToList();
        }

        /// <summary>
        /// Initializes the redmine state.
        /// </summary>
        public void InitializeRedmineNew()
        {
            var t = new Thread(this.InitializeRedmineThreadMethod);

            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        /// <summary>
        /// Method to try to reload an issue by its id
        /// </summary>
        /// <param name="issueId">the issue id</param>
        /// <returns>the corresponding issue info and the list of new issues</returns>
        public Tuple<IssueInfo, List<IssueInfo>> ReloadIssueById(int issueId)
        {
            IssueInfo issueInfo = null;
            var newIssueList = new List<IssueInfo>();

            if (!this.AllIssues.TryGetValue(issueId, out issueInfo))
            {
                // we do not know the issue yet, try to reload it
                try
                {
                    // try to reload the ticket by its ticket number
                    var parameter = new DataSourceParameter() { IssueId = issueId };
                    var issues = this._externalDataSource.GetIssueInfoList(parameter);
                    issueInfo = issues.FirstOrDefault();

                    // it ticket was found, add it to list of all issues and update the known issues
                    if (issueInfo != null)
                    {
                        this.AllIssues.Add(issueId, issueInfo);
                        var issueList = this.AllIssues.Values.Distinct().ToList();
                        LocalCache.WriteObject(LocalCache.KnownIssues, issueList);
                        newIssueList.Add(issueInfo);
                    }
                    else
                    {
                        // if the ticket could not be found by its ticket number, it probably has status done and cannot be found via the api
                        // hence the usual ticket reloading is triggered, which updates all tickets from redmine

                        // start reloading of issues
                        var reloadedIssues = DownloadHelper.ReloadIssueInfoExtended(issueId, this._externalDataSource, this.AllIssues.Values.ToList());

                        // get all new issues and add them to the issue list
                        var newIssues = reloadedIssues.Keys.Except(this.AllIssues.Keys).ToList();
                        reloadedIssues.Where(i => newIssues.Contains(i.Key)).ForEach(i => this.AllIssues.Add(i.Key, i.Value));
                        
                        // update the known issue list
                        var issueList = this.AllIssues.Values.Distinct().ToList();
                        LocalCache.WriteObject(LocalCache.KnownIssues, issueList);

                        // look for the missing ticket
                        issueInfo = reloadedIssues.Values.FirstOrDefault(i => object.Equals(i.Id, issueId));
                        var infoText = string.Format(
                            "No result for loading ticket #{0}, reloading all tickets ({1} new tickets found).",
                            issueId,
                            newIssues.Count());
                        Log.Info(infoText);
                        newIssueList.AddRange(reloadedIssues.Values.Where(i => newIssues.Contains(i.Id.GetValueOrDefault(-1))));
                    }

                    // if still no ticket can be found, log a warning 
                    if (issueInfo == null)
                    {
                        Log.Warn(string.Format("No result for loading ticket #{0}, {1} results.", issueId, issues.Count));
                    }
                }
                catch (Exception exception)
                {
                    var text = string.Format("Error while reloading issue with id {0}", issueId);
                    Log.Error(text, exception);
                }
            }

            return Tuple.Create(issueInfo, newIssueList);
        }

        /// <summary>
        /// Forces synchronization of the current view
        /// </summary>
        /// <returns>
        /// An awaitable task.
        /// </returns>
        public async Task RevertTimeEntriesToRedmineState()
        {
            if (!this.CanSyncTimeEntries)
            {
                return;
            }

            var modifiedAppointments = this.Calendar.GetAppointmentsWithModification();

            var earliestModified = DateTime.Now.Date;
            var latestModified = DateTime.Now.Date;

            if (modifiedAppointments != null && modifiedAppointments.Any())
            {
                earliestModified = modifiedAppointments.Min(app => app.Start);
                latestModified = modifiedAppointments.Max(app => app.End);
            }

            var wnd = new OfficeWin32Window(Globals.ThisAddIn.Application.ActiveWindow());
            var dialog = new RevertDialog(earliestModified, latestModified);
            var wih = new WindowInteropHelper(dialog) { Owner = wnd.Handle };
            dialog.ShowDialog();

            if (!dialog.DialogResult.HasValue || !dialog.DialogResult.Value)
            {
                // The user canceled, do nothing.
                return;
            }

            this.CanSyncTimeEntries = false;
            Globals.ThisAddIn.SyncState.RaiseConnectionChanged();
            Globals.ThisAddIn.SyncState.Status = "Setze Zeiteinträge zurück...";

            // wait for the sync completion
            await Task.Run(() => this.ResetTimeEntries(dialog.StartDate, dialog.EndDate));

            this.CanSyncTimeEntries = true;
            Globals.ThisAddIn.SyncState.Status = "Verbunden";
            Globals.ThisAddIn.SyncState.RaiseConnectionChanged();

            this.RaiseAppointmentChanged();
        }

        /// <summary>
        /// Saves the time entries asynchronously.
        /// </summary>
        /// <returns>An awaitable task.</returns>
        public async Task SaveTimeEntriesAsync()
        {
            if (!this.CanSyncTimeEntries)
            {
                return;
            }

            var modifiedAppointments = this.Calendar.GetAppointmentsWithModification();

            // When there are deleted entries, ask the user for confirmation to delete.
            if (modifiedAppointments.Any(i => i.IsDeletedSet()))
            {
                var wnd = new OfficeWin32Window(Globals.ThisAddIn.Application.ActiveWindow());
                var dialog = new SaveDialog(modifiedAppointments);
                var wih = new WindowInteropHelper(dialog) { Owner = wnd.Handle };
                dialog.ShowDialog();

                if (!dialog.DialogResult.GetValueOrDefault(false))
                {
                    // The user canceled, do nothing.
                    return;
                }
            }
            this.CanSyncTimeEntries = false;
            Globals.ThisAddIn.SyncState.Status = "Speichere Änderungen...";
            Globals.ThisAddIn.SyncState.RaiseConnectionChanged();

            await Task.Run(() => this.SaveTimeEntries(modifiedAppointments));

            this.CanSyncTimeEntries = true;
            Globals.ThisAddIn.SyncState.Status = "Verbunden";
            Globals.ThisAddIn.SyncState.RaiseConnectionChanged();
        }

        /// <summary>
        /// Updates the issue for an appointment, is called after the creation of an appointment or on change of it.
        /// </summary>
        /// <param name="appointment">The appointment item</param>
        /// <param name="issueId">The new issue id</param>
        public void UpdateAppointmentIssue(AppointmentItem appointment, int issueId)
        {
            var originalId = appointment.GetIssueId();

            // if the the id did not change, or the new issue number is not know, skip further processing.
            if (originalId == issueId || !this._issues.ContainsKey(issueId))
            {
                return;
            }

            // set issue to appointment
            var issue = this._issues[issueId];
            appointment.SetProjectId(issue.ProjectId);
            appointment.SetIssueId(issueId);
            appointment.CreateAppointmentLocation(issueId, issue);
            appointment.SetAppointmentState(AppointmentState.Modified);
            appointment.Save();

            // update last used issues
            var issueRef = this.AllIssues[issueId];
            this.LastUsedIssues.Remove(issueRef);
            this.LastUsedIssues.Insert(0, issueRef);
            if (this.LastUsedIssues.Count > this._numberLastUsedIssues)
            {
                this.LastUsedIssues = this.LastUsedIssues.Take(this._numberLastUsedIssues).ToList();
            }
            Settings.Default.LastUsedIssues = string.Join(";", this.LastUsedIssues.Select(iref => iref.Id));
            Settings.Default.Save();
            this.RaiseAppointmentChanged();
        }

        /// <summary>
        /// Saves the favorite issues to the settings.
        /// </summary>
        /// <param name="issues">The issue information of the issues which shall become the new favorite issues.</param>
        public void UpdateFavoriteIssues(List<IssueInfo> issues)
        {
            this.FavoriteIssues = issues;

            Settings.Default.FavoriteIssues = string.Join(";", this.FavoriteIssues.Select(iref => iref.Id));
            Settings.Default.Save();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Method that creates recurring time entries by showing a dialog to the user for providing the necessary information.
        /// </summary>
        internal void CreateRecurringTimeEntries()
        {
            var wnd = new OfficeWin32Window(Globals.ThisAddIn.Application.ActiveWindow());
            var dialog = new RecurringTimeEntryDialog();
            var wih = new WindowInteropHelper(dialog) { Owner = wnd.Handle };
            dialog.AvailableIssues = this.AllIssues.Values.ToList();
            dialog.ShowDialog();

            if (dialog.DialogResult.HasValue && dialog.DialogResult.Value)
            {
                var issue = this._issues[dialog.SelectedIssue.Id.Value];
                var project = this._projects[dialog.SelectedIssue.ProjectId];
                var activity = this.GetDefaultActivity();

                var currentDate = dialog.StartDate.Date;
                while (currentDate <= dialog.EndDate.Date)
                {
                    // Only create entries for workdays, except the user wanted to book on weekends.
                    if (DateTimeHelper.IsWorkDay(currentDate) || dialog.IsBookingOnWeekends)
                    {
                        // create an appointment for this day
                        try
                        {
                            // create & save the appointment
                            var newEntry = (AppointmentItem)this.Calendar.Items.Add(OlItemType.olAppointmentItem);
                            newEntry.SetIsImported(false);
                            newEntry.ReminderSet = false;

                            newEntry.CreateAppointmentLocation(issue.Id.Value, issue);
                            newEntry.Subject = dialog.Description;
                            newEntry.Start = currentDate.Date.AddHours(dialog.StartTime.Hour).AddMinutes(dialog.StartTime.Minute);
                            newEntry.End = currentDate.Date.AddHours(dialog.EndTime.Hour).AddMinutes(dialog.EndTime.Minute);

                            newEntry.UpdateAppointmentFields(null, project.Id.Value, issue.Id.Value, activity.Id.Value, DateTime.Now);
                            newEntry.SetAppointmentState(AppointmentState.Modified);

                            newEntry.Save();
                            newEntry.MarkAsNotCopied();

                            // register for change tracking
                            this.RegisterAppointment(newEntry);
                        }
                        catch (Exception ex)
                        {
                            Log.Error("Could not create new appointments in redmine calendar. ", ex);
                        }
                    }
                    currentDate = currentDate.AddDays(1);
                }
            }
        }

        
        /// <summary>
        /// Creates a time entry for an appointment in redmine
        /// </summary>
        /// <param name="item">
        /// The item to transfer
        /// </param>
        private void CreateInExternalSource(AppointmentItem item)
        {
            var updateTime = DateTime.Now;
            var entry = this.CreateTimeEntryFromAppointment(item, updateTime);
            if (entry != null)
            {
                entry.Id = 0;
                var resultEntry = this._externalDataSource.CreateObject(entry);

                // update the appointment properties
                item.SetTimeEntryId(resultEntry.Id);
                if (resultEntry.IssueInfo.Id != Settings.Default.RedmineUseOvertimeIssue)
                {
                    item.SetAppointmentState(AppointmentState.Synchronized);
                }
                else
                {
                    item.SetAppointmentState(AppointmentState.SynchronizedOvertime);
                }
                item.SetAppointmentModificationDate(resultEntry.UpdateTime);
                item.Save();
                this.UpdateCache(resultEntry);
            }
        }

        /// <summary>
        /// Creates a time entry for an appointment
        /// </summary>
        /// <param name="item">
        /// The appointment item
        /// </param>
        /// <param name="updateTime">
        /// The update Time.
        /// </param>
        /// <returns>
        /// The redmine time entry
        /// </returns>
        private TimeEntryInfo CreateTimeEntryFromAppointment(AppointmentItem item, DateTime updateTime)
        {
            var entryId = item.GetTimeEntryId();
            var projectId = item.GetProjectId();
            var issueId = item.GetIssueId();
            var activityId = item.GetActivityId();

            // mandatory fields have to be set
            if (activityId == null)
            {
                // get default activity
                var defaultAct = this.GetDefaultActivity();

                // update appointment
                item.SetActivityId(defaultAct.Id);
                item.Save();

                activityId = defaultAct.Id;
            }

            // check if there is an issue
            if (issueId == null)
            {
                // try to identify issue based on appointment subject
                var issueGuess = this.TryGuessIssue(item);
                if (issueGuess != null)
                {
                    item.SetIssueId(issueGuess.Id);
                    item.Save();

                    issueId = issueGuess.Id;
                }
                else
                {
                    item.AppendToBody("Das Issue für diesen Eintrag konnte nicht ermittelt werden.");
                    item.SetAppointmentState(AppointmentState.SyncError);
                    item.Save();
                    return null;
                }
            }
            else
            {
                // check if the saved issue id is ok
                if (!this._issues.ContainsKey(issueId.Value))
                {
                    item.AppendToBody("Das Issue für diesen Eintrag konnte nicht ermittelt werden.");
                    item.SetAppointmentState(AppointmentState.SyncError);
                    item.Save();
                    return null;
                }
            }

            // we have a issue number but no project number
            if (projectId == null)
            {
                var issue = this._issues[issueId.Value];
                item.SetProjectId(issue.ProjectId);
                item.Save();
                projectId = issue.ProjectId;
            }

            // get the corresponding info objects
            ActivityInfo activityInfo = null;
            IssueInfo issueInfo = null;
            ProjectInfo projectInfo = null;

            this._activities.TryGetValue(activityId.Value, out activityInfo);
            this._issues.TryGetValue(issueId.Value, out issueInfo);
            this._projects.TryGetValue(projectId.Value, out projectInfo);

            TimeEntryInfo timeEntryInfoToCreate = null;
            if (activityInfo != null && issueInfo != null && projectInfo != null)
            {
                // create the time entry info if possible
                timeEntryInfoToCreate = new TimeEntryInfo()
                                            {
                                                Id = entryId,
                                                StartDateTime = item.Start,
                                                EndDateTime = item.End,
                                                Name = item.Subject,
                                                UpdateTime = updateTime,
                                                ActivityInfo = activityInfo,
                                                IssueInfo = issueInfo,
                                                ProjectInfo = projectInfo,
                                            };
            }
            else
            {
                item.AppendToBody("Das Issue, das Projekt oder die Aktivität für diesen Eintrag konnten nicht ermittelt werden.");
                item.SetAppointmentState(AppointmentState.SyncError);
                item.Save();
                return null;
            }

            return timeEntryInfoToCreate;
        }

        /// <summary>
        /// Deletes the time entry that is associated with an <see cref="AppointmentItem"/> from redmine.
        /// </summary>
        /// <param name="item">The item for which to delete the corresponding time entry from redmine.</param>
        /// <returns>if deletion was successful</returns>
        private bool DeleteTimeEntryInfo(AppointmentItem item)
        {
            var timeEntryId = item.GetTimeEntryId();
            if (!timeEntryId.HasValue)
            {
                // if the item has no time entry id, it is not in redmine, thus consider the deletion successful.
                return true;
            }

            try
            {
                // call the external source to delete the entry
                this._externalDataSource.DeleteTimeEntry(timeEntryId.Value, new DataSourceParameter());
                return true;
            }
            catch (ConnectionException connectionException)
            {
                Log.Error(connectionException.Message);
                throw;
            }
        }

        
        

        /// <summary>
        /// Gets the default activity for time entries.
        /// </summary>
        /// <returns>The default activity for time entries.</returns>
        private ActivityInfo GetDefaultActivity()
        {
            var defaultAct = this._activities.Select(act => act.Value).FirstOrDefault(act => act.IsDefault);
            if (defaultAct == null)
            {
                defaultAct = this._activities.Select(act => act.Value).First();
            }

            return defaultAct;
        }

        /// <summary>
        /// Gets the issue for a time entry if the issue has been loaded.
        /// </summary>
        /// <param name="entry">The time entry for which to get the issue.</param>
        /// <returns>The corresponding issue for the time entry. <code>null</code> if there is no issue loaded.</returns>
        private IssueInfo GetIssueForTimeEntry(TimeEntryInfo entry)
        {
            return entry.IssueInfo;
        }

        /// <summary>
        /// Method which is executed by the synchronization thread.
        /// </summary>
        private void InitializeRedmineThreadMethod()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(Settings.Default.RedmineApiKey))
                {
                    MessageBox.Show("Es ist kein API-Key angegeben, eine Verbindung zu Redmine kann nicht hergestellt werden.");
                    throw new Exception("Could not connect to redmine, because the API-key is missing.");
                }

                this.IsConnecting = true;
                this.CanSyncTimeEntries = false;

                Globals.ThisAddIn.SyncState.Status = "Verbinde...";
                Globals.ThisAddIn.SyncState.RaiseConnectionChanged();

                // Initialize the redmine manager
                this._externalDataSource = ExternalDataSourceFactory.GetRedmineMangerInstance(
                    this.ConnectionUrl,
                    Settings.Default.RedmineApiKey,
                    Settings.Default.LimitForIssueNumber);

                // Get user information
                this.CurrentUser = this._externalDataSource.GetCurrentUser();

                // get activity infos
                this._activities = DownloadHelper.GetActivityInfos(this._externalDataSource);

                // load projects
                var projectLists = DownloadHelper.GetCurrentProjectsAndUpdateCache(this._externalDataSource);
                var newProjects = projectLists.NewProjects;
                this._projects = projectLists.AllProjects;

                // load issues
                this._issues = DownloadHelper.DownloadIssues(newProjects, this._issues, this._externalDataSource);
                
                // update the special issue infos
                this.UpdateSpecialIssueInformation();

                // update the settings and display correct state
                Globals.ThisAddIn.SyncState.Status = "Verbunden";
                this.CanSyncTimeEntries = true;
                this.IsConnecting = false;
                Globals.ThisAddIn.SyncState.RaiseConnectionChanged();
                this.RaiseAppointmentChanged();
            }
            catch (Exception ex)
            {
                Log.Error("Error in Initialize.", ex);
                this.CurrentUser = null;
                this.CanSyncTimeEntries = false;
                Globals.ThisAddIn.SyncState.Status = "Nicht verbunden oder Fehler bei Initialisierung";
            }
            finally
            {
                this.IsConnecting = false;
            }
        }

        /// <summary>
        /// Method called to update the special issue infos
        /// </summary>
        private void UpdateSpecialIssueInformation()
        {
            // update favorite and last used issues
            this.SetSpecialIssues(Settings.Default.LastUsedIssues, this.LastUsedIssues);
            this.SetSpecialIssues(Settings.Default.FavoriteIssues, this.FavoriteIssues);

            // update the last used issues
            this.UpdateLastUsedIssues();
        }

        /// <summary>
        /// Is invoked before an appointment item is deleted
        /// </summary>
        /// <param name="item">The item</param>
        /// <param name="cancel">Cancel deletion?</param>
        private void ItemOnBeforeDelete(object item, ref bool cancel)
        {
            var appItem = (AppointmentItem)item;

            var redmineId = appItem.GetTimeEntryId();
            if (redmineId != null && redmineId.Value > 0)
            {
                // If the appointment has a corresponding time entry in redmine, do not delete it immediately.
                cancel = true;
                appItem.SetAppointmentState(AppointmentState.Deleted);
                appItem.Save();
            }
        }

        /// <summary>
        /// Is invoked when a property of an appointment changed
        /// </summary>
        /// <param name="item">The item</param>
        /// <param name="name">The property name that changed</param>
        private void ItemOnPropertyChange(AppointmentItem item, string name)
        {
            // get change reason
            var subjectChanged = name == "Subject";
            var timesChanged = name == "Start" || name == "End";
            var hasIssueInfo = item.GetIssueId().HasValue || this.TryGuessIssue(item) != null;

            if ((subjectChanged || timesChanged) && hasIssueInfo)
            {
                if (!item.IsModifiedSet())
                {
                    // we only need to set the item to state modified, if it was not before
                    item.SetAppointmentState(AppointmentState.Modified);
                    item.Save();
                }

                // the raise appointment update method is called here to update the ui
                if (timesChanged)
                {
                    this.RaiseAppointmentChanged();
                }
            }
        }

        /// <summary>
        /// Callback method that is used whenever a new appointment is added to the redmine calendar.
        /// </summary>
        /// <param name="item">The new item that was added to the calendar.</param>
        private void OnItemsOnItemAdd(object item)
        {
            var appItem = item as AppointmentItem;
            if (appItem == null)
            {
                return;
            }

            appItem.ReminderSet = false;

            /* 
             * We need to check if the appointment is copied. We could do this with appItem.IsAppointmentCopied(). However, 
             * this does not work because that method relies on the id of the appointment given by outlook, which is only set
             * when the appointment is saved. Therefore, set the Appointmentstate to modified, when the appointment that is 
             * added already has the Field Constants.FieldRedmineIssueId set (i.e. it already has an issue to which it belongs).
             * However, this would also set appointments as changed that were imported from redmine. Therefore the redmine 
             * synchronizer sets another custom property: IsImported. This indicates that we must not change the state to modified here.
             */
            var isImported = appItem.IsImported();
            if ((appItem.GetIssueId().HasValue && !isImported)
                || (!appItem.GetIssueId().HasValue && this.TryGuessIssue(appItem) != null))
            {
                appItem.SetAppointmentState(AppointmentState.Modified);
            }
            else
            {
                appItem.ClearIsImported();
            }
            appItem.Save();
            this.RegisterAppointment(appItem);
        }

        /// <summary>
        /// Raises the appointment changed event
        /// </summary>
        private void RaiseAppointmentChanged()
        {
            if (this.AppointmentChanged != null)
            {
                this.AppointmentChanged(this, new EventArgs());
            }
        }

        /// <summary>
        /// Registers an appointment for change tracking
        /// </summary>
        /// <param name="item">The item to track</param>
        private void RegisterAppointment(AppointmentItem item)
        {
            lock (this._managedItems)
            {
                if (this._managedItems.Contains(item))
                {
                    return;
                }

                item.BeforeDelete += this.ItemOnBeforeDelete;
                item.PropertyChange += (name) => this.ItemOnPropertyChange(item, name);
                this._managedItems.Add(item);
            }
        }

        /// <summary>
        /// Reverts the time entries that are on and between the start and end date.
        /// </summary>
        /// <param name="startDate">The start date for reverting.</param>
        /// <param name="endDate">The end date for reverting.</param>
        private void ResetTimeEntries(DateTime startDate, DateTime endDate)
        {
            // query redmine entries for the time range
            var parameters = new DataSourceParameter() { UseLimit = true, UserId = -1, SpentDateTimeTuple = Tuple.Create(startDate, endDate) };

            var retrievedItems = new List<TimeEntryInfo>();
            try
            {
                var items = this._externalDataSource.GetTotalTimeEntryInfoList(
                    parameters,
                    (cur, total) => Globals.ThisAddIn.SyncState.Status = string.Format("Lade Zeiteinträge ({0}/{1})", Math.Min(cur, total), total));
                retrievedItems = items.ToList();
            }
            catch (Exception ex)
            {
                Log.Error("Exception while retrieving time entries from redmine.", ex);
                MessageBox.Show(
                    "Die Zeiteinträge konnten nicht von Redmine abgerufen werden. Eventuell besteht keine Internetverbindung.\n\n" + ex.ToString());
                return;
            }

            try
            {
                // get the outlook items
                var itemsForDateRange = this.Calendar.GetAppointmentsInRange(startDate.Date, endDate.Date.AddDays(1).AddSeconds(-1));

                var entries = retrievedItems.ToDictionary(te => te.Id, te => te);

                // check existing items
                foreach (var appointment in itemsForDateRange)
                {
                    // already synced item?
                    var redmineId = appointment.GetTimeEntryId();
                    if (redmineId == null || appointment.IsAppointmentCopied())
                    {
                        // delete items that does not exist in redmine
                        appointment.SetTimeEntryId(null);
                        appointment.Save();
                        appointment.Delete();
                    }
                    else
                    {
                        // this is an existing item that should have a redmine time entry
                        var rid = Convert.ToInt32(redmineId.Value);
                        if (entries.ContainsKey(rid))
                        {
                            var redmineEntry = entries[rid];

                            if (appointment.CheckItemIsModified(redmineEntry))
                            {
                                // update appointment
                                appointment.UpdateAppointmentFromTimeEntry(redmineEntry, this.GetIssueForTimeEntry(redmineEntry));
                                appointment.SetAppointmentModificationDate(redmineEntry.UpdateTime);
                            }

                            if (redmineEntry.IssueInfo.Id != Settings.Default.RedmineUseOvertimeIssue)
                            {
                                appointment.SetAppointmentState(AppointmentState.Synchronized);
                            }
                            else
                            {
                                appointment.SetAppointmentState(AppointmentState.SynchronizedOvertime);
                            }
                            appointment.Save();

                            entries.Remove(rid);
                        }
                        else
                        {
                            // there is no corresponding redmine item although this item has a redmine id
                            // it might be deleted in redmine so it is now deleted here.
                            // first set the redmine id to null so that the delete function will not mark it again as deleted in redmine
                            appointment.SetTimeEntryId(null);
                            appointment.Delete();
                        }
                    }
                }

                // process remaining (=new in redmine) items
                foreach (var entry in entries)
                {
                    #region Create appointment for time entry

                    // create & save the appointment
                    var timeEntry = entry.Value;
                    var newEntry = (AppointmentItem)this.Calendar.Items.Add(OlItemType.olAppointmentItem);
                    newEntry.SetIsImported(true);
                    newEntry.ReminderSet = false;
                    newEntry.UpdateAppointmentFromTimeEntry(timeEntry, this.GetIssueForTimeEntry(timeEntry));
                    newEntry.Save();
                    newEntry.MarkAsNotCopied();

                    // register for change tracking
                    this.RegisterAppointment(newEntry);

                    #endregion
                }
            }
            catch (Exception ex)
            {
                Debugger.Break();
                Log.Error("Error while handling time entries.", ex);
                MessageBox.Show("Beim Verarbeiten von Zeiteinträgen ist ein Fehler aufgetreten:\n\n" + ex.ToString());
            }
        }

        /// <summary>
        /// Saves the time entries to redmine which are provided as a parameter.
        /// </summary>
        /// <param name="modified">The time entries to save.</param>
        private void SaveTimeEntries(List<AppointmentItem> modified)
        {
            var errors = new List<string>();
            var totalModified = modified.Count();
            var currentItem = 0;

            foreach (var item in modified)
            {
                try
                {
                    currentItem++;
                    Globals.ThisAddIn.SyncState.Status = string.Format("Speichere {0} von {1}", currentItem, totalModified);

                    // If the item was conflicted, try to reset its previous state, 
                    // modified or deleted, s.t. we can do another synchronization.
                    if (item.IsSyncErrorSet())
                    {
                        item.ResetPreviousState();
                    }

                    this.SaveTimeEntry(item);
                }
                catch (ConnectionException connectionException)
                {
                    // set correct state
                    var previousAppointmentState = connectionException.PreviousAppointmentState;
                    item.SetAppointmentState(previousAppointmentState);
                    item.Save();

                    // add error message
                    var message = string.Format(
                        "Aufgrund fehlender Verbindung konnte Zeiteintrag {0} konnte nicht synchronisiert werden.",
                        item.GetStringRepresentation());
                    errors.Add(message);
                    Log.Error(message, connectionException);
                }
                catch (Exception ex)
                {
                    // put the exception details to the appointment body
                    item.AppendToBody(ex.ToString());
                    item.SetAppointmentState(AppointmentState.SyncError);
                    item.Save();

                    var message = string.Format("Zeiteintrag {0} konnte nicht synchronisiert werden.", item.GetStringRepresentation());
                    errors.Add(message);
                    Log.Error(message, ex);
                }
            }

            this.RaiseAppointmentChanged();
            Globals.ThisAddIn.SyncState.RaiseConnectionChanged();
            if (errors.Any())
            {
                MessageBox.Show(string.Join("\n", errors), "Fehler bei der Synchronisation");
            }
        }

        /// <summary>
        /// Saves a single time entry to redmine.
        /// </summary>
        /// <param name="item">The appointment item which corresponds to the time entry which shall be saved.</param>
        private void SaveTimeEntry(AppointmentItem item)
        {
            try
            {
                if (item.IsModifiedSet())
                {
                    var timeEntryId = item.GetTimeEntryId();
                    if (timeEntryId.HasValue && !item.IsAppointmentCopied())
                    {
                        // update an existing appointment
                        this.UpdateInExternalSource(item);
                    }
                    else
                    {
                        // create a new appointment
                        this.CreateInExternalSource(item);
                        item.MarkAsNotCopied();
                    }
                }
                else if (item.IsDeletedSet())
                {
                    // delete an appointment
                    var success = this.DeleteTimeEntryInfo(item);

                    if (success)
                    {
                        item.SetTimeEntryId(null);
                        item.Delete();
                    }
                }
                else
                {
                    // we should never get here, since we only get appointments modified in one way or another
                    Log.Warn(string.Format("Item is assumed modified, but has no category set for modified or deleted: {0}", item.Location));
                }
            }
            catch (ConnectionException connectionException)
            {
                // add info to previous state of the appointment
                connectionException.PreviousAppointmentState = item.IsModifiedSet() ? AppointmentState.Modified : AppointmentState.Deleted;
                throw;
            }
        }

        /// <summary>
        /// Method to set special issues
        /// </summary>
        /// <param name="specialIssueIds">the ids of the special issues, concated using ";"</param>
        /// <param name="issueList">the issue list to store the issues in</param>
        private void SetSpecialIssues(string specialIssueIds, List<IssueInfo> issueList)
        {
            // create last used issues
            if (!string.IsNullOrWhiteSpace(specialIssueIds))
            {
                issueList.Clear();
                foreach (var item in specialIssueIds.Split(';'))
                {
                    try
                    {
                        var issueId = Convert.ToInt32(item);
                        IssueInfo issueInfo;
                        if (this.AllIssues.TryGetValue(issueId, out issueInfo))
                        {
                            issueList.Add(issueInfo);
                        }
                    }
                    catch (Exception ex)
                    {
                        var msg = string.Format("Could not add issue with id {0} to list of special issues.", item);
                        Log.Error(msg, ex);
                    }
                }
            }
        }

        /// <summary>
        /// Tries to guess the correct issue from the appointment subject line
        /// </summary>
        /// <param name="item">The appointment item</param>
        /// <returns>The issue or null if no issue could be found</returns>
        private IssueInfo TryGuessIssue(AppointmentItem item)
        {
            if (RxIssueNumber.IsMatch(item.Subject))
            {
                try
                {
                    var m = RxIssueNumber.Match(item.Subject);
                    var issueIdStr = m.Groups["IssueNumber"].Captures[0].Value;
                    var id = Convert.ToInt32(issueIdStr);

                    // we get the issue from the existing issues or try to reload it
                    return this.ReloadIssueById(id).Item1;
                }
                catch (Exception ex)
                {
                    Log.Error("Could not get Issue", ex);
                    return null;
                }
            }
            return null;
        }

        /// <summary>
        /// Updates a time entry in redmine from an appointment
        /// </summary>
        /// <param name="item">
        /// The item to transfer
        /// </param>
        private void UpdateInExternalSource(AppointmentItem item)
        {
            var updateTime = DateTime.Now;
            var timeEntryInfo = this.CreateTimeEntryFromAppointment(item, updateTime);
            if (timeEntryInfo != null)
            {
                // update object in external source
                var resultEntry = this._externalDataSource.UpdateObject(timeEntryInfo);

                item.SetTimeEntryId(resultEntry.Id);
                if (resultEntry.IssueInfo.Id != Settings.Default.RedmineUseOvertimeIssue)
                {
                    item.SetAppointmentState(AppointmentState.Synchronized);
                }
                else
                {
                    item.SetAppointmentState(AppointmentState.SynchronizedOvertime);
                }

                item.SetAppointmentModificationDate(resultEntry.UpdateTime);
                item.Save();
                this.UpdateCache(resultEntry);
            }
        }

        /// <summary>
        /// Updates the cached project and issue info.
        /// </summary>
        /// <param name="resultEntry">
        /// The time entry info
        /// </param>
        private void UpdateCache(TimeEntryInfo resultEntry)
        {
            if (resultEntry.ProjectInfo.Id.HasValue)
            {
                this._projects[resultEntry.ProjectInfo.Id.Value] = resultEntry.ProjectInfo;
            }
            if (resultEntry.IssueInfo.Id.HasValue)
            {
                this._issues[resultEntry.IssueInfo.Id.Value] = resultEntry.IssueInfo;
            }
        }

        /// <summary>
        /// Method to update the list of the issues last used
        /// </summary>
        private void UpdateLastUsedIssues()
        {
            // set number of last used issues to use
            this._numberLastUsedIssues = Math.Max(0, Settings.Default.NumberLastUsedIssues);

            if (this.LastUsedIssues.Count > this._numberLastUsedIssues)
            {
                var elementsToRemain = new HashSet<IssueInfo>(this.LastUsedIssues.Take(this._numberLastUsedIssues));
                this.LastUsedIssues.RemoveAll(i => !elementsToRemain.Contains(i));
            }

            var lastUsedIssues = Settings.Default.LastUsedIssues;
            var split = lastUsedIssues.Split(';');
            if (split.Length > this._numberLastUsedIssues)
            {
                var lastUsedToTake = split.Take(this._numberLastUsedIssues);
                var lengthOfNew = lastUsedToTake.Sum(e => e.Length) + this._numberLastUsedIssues - 1;
                var newLastUsed = lastUsedIssues.Substring(0, lengthOfNew);
                Settings.Default.LastUsedIssues = newLastUsed;
                Settings.Default.Save();
            }
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Gets the user which is currently logged into redmine.
        /// </summary>
        public UserInfo CurrentUser { get; private set; }

        /// <summary>
        /// Gets the reference to the redmine calendar
        /// </summary>
        public MAPIFolder Calendar { get; private set; }

        /// <summary>
        /// Gets the items collection of the redmine calendar. This reference has to be preserved 
        /// in order to keep the ItemAdd listener on the item collection.
        /// </summary>
        public Items CalendarItems { get; private set; }

        /// <summary>
        /// Gets a value indicating whether there is a connection to the redmine system
        /// </summary>
        public bool IsConnected
        {
            get
            {
                return this._externalDataSource != null && this.CurrentUser != null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether or not a synchronization of projects and issues is in progress.
        /// </summary>
        public bool IsConnecting { get; private set; }

        /// <summary>
        /// Gets a value indicating whether a forced sync with save is possible at the moment
        /// </summary>
        public bool CanSyncTimeEntries { get; private set; }

        /// <summary>
        /// Gets a list of all issues
        /// </summary>
        public IDictionary<int, IssueInfo> AllIssues
        {
            get
            {
                return this._issues;
            }
        }

        /// <summary>
        /// Gets a list of previously used issues
        /// </summary>
        public List<IssueInfo> LastUsedIssues { get; private set; }

        /// <summary>
        /// Gets a list of issues that were marked as favorite issues.
        /// </summary>
        public List<IssueInfo> FavoriteIssues { get; private set; }

        /// <summary>
        /// Gets the URL to the external system
        /// </summary>
        public string ConnectionUrl
        {
            get
            {
                return Settings.Default.RedmineURL.Last() == '/'
                           ? Settings.Default.RedmineURL.Substring(0, Settings.Default.RedmineURL.Length - 1)
                           : Settings.Default.RedmineURL;
            }
        }

        /// <summary>
        /// Gets the current user name.
        /// </summary>
        public string CurrentUserName
        {
            get
            {
                return this.CurrentUser != null ? this.CurrentUser.Name : string.Empty;
            }
        }

        #endregion
    }
}