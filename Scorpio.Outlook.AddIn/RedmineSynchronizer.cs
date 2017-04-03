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

namespace Scorpio.Outlook.AddIn
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Diagnostics;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Timers;
    using System.Windows.Forms;
    using System.Windows.Interop;

    using log4net;

    using Microsoft.Office.Interop.Outlook;

    using Redmine.Net.Api;
    using Redmine.Net.Api.Types;

    using Scorpio.Outlook.AddIn.Extensions;
    using Scorpio.Outlook.AddIn.Misc;
    using Scorpio.Outlook.AddIn.Properties;
    using Scorpio.Outlook.AddIn.UserInterface.Controls;

    using Exception = System.Exception;
    using Timer = System.Timers.Timer;

    /// <summary>
    /// Class that is responsible for keeping redmine time entries in sync with the redmine calendar in outlook.
    /// </summary>
    public class RedmineSynchronizer
    {
        #region Static Fields

        /// <summary>
        /// The logger.
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(RedmineSynchronizer));

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
        /// List of all AppointmentItems in the redmine calendar. Elements in this set have a deletionlistener and a changelistener attached.
        /// </summary>
        private readonly HashSet<AppointmentItem> _managedItems = new HashSet<AppointmentItem>();

        /// <summary>
        /// This timer is used for keeping ui elements up to date.
        /// </summary>
        private readonly Timer _updateTimer;

        /// <summary>
        /// A dictionary of all time entries that the user can read in redmine. Mapping from time entry id to time entry.
        /// </summary>
        private IDictionary<int, TimeEntryActivity> _activities = new Dictionary<int, TimeEntryActivity>();

        /// <summary>
        /// A dictionary of all issues that the user can read in redmine. Mapping from issue id to issue.
        /// </summary>
        private IDictionary<int, Issue> _issues = new Dictionary<int, Issue>();

        /// <summary>
        /// A dictionary of all projects that the user can read in redmine. Mapping from project id to project.
        /// </summary>
        private IDictionary<int, Project> _projects = new Dictionary<int, Project>();

        /// <summary>
        /// The redmine manager, an API wrapper for the redmine rest api.
        /// </summary>
        private RedmineManager _redmineManager;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="RedmineSynchronizer"/> class.
        /// </summary>
        /// <param name="calendar">The target calendar in outlook for syncing redmine time entries</param>
        public RedmineSynchronizer(MAPIFolder calendar)
        {
            this.Calendar = calendar;
            this._updateTimer = new Timer(Constants.TimerInterval);
            this._updateTimer.Elapsed += this.UpdateTimerCallback;
            this._updateTimer.AutoReset = true;
            this._updateTimer.SynchronizingObject = new Control();
            this._updateTimer.Enabled = true;
            Globals.ThisAddIn.SyncState.Status = "Nicht verbunden";
            this.LastUsedIssues = new List<IssueProjectInfo>();
            this.FavoriteIssues = new List<IssueProjectInfo>();
            this.AllIssues = new Dictionary<int, IssueProjectInfo>();

            // At startup, register all appointments in the redmine calendar for change tracking.
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
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Build the issue information string (link + name)
        /// </summary>
        /// <param name="issueId">The issue id</param>
        /// <returns>The link and the name of the issue</returns>
        public Tuple<string, string> BuildIssueInformation(int issueId)
        {
            var issueLink = string.Format("{0}/issues/{1}", this.RedmineUrl, issueId);
            var issueName = string.Format("Issue {0}", issueId);

            if (this._issues.ContainsKey(issueId))
            {
                issueName = this._issues[issueId].Subject;
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
            var projectLink = string.Format("{0}/projects/{1}", this.RedmineUrl, projectId);
            var projectName = string.Format("Project {0}", projectId);

            if (this._projects.ContainsKey(projectId))
            {
                projectName = this._projects[projectId].Name;
            }
            return new Tuple<string, string>(projectLink, projectName);
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

                if (!dialog.DialogResult.HasValue || !dialog.DialogResult.Value)
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
        /// Updates the issue for an appointment
        /// </summary>
        /// <param name="appointment">The appointment item</param>
        /// <param name="issueId">The new issue id</param>
        public void UpdateAppointmentIssue(AppointmentItem appointment, int issueId)
        {
            var originalId = appointment.GetAppointmentCustomId(Constants.FieldRedmineIssueId);

            // if the the id did not change, or the new issue number is not know, skip further processing.
            if (originalId == issueId || !this._issues.ContainsKey(issueId))
            {
                return;
            }

            // set issue to appointment
            var issue = this._issues[issueId];
            appointment.SetAppointmentCustomId(Constants.FieldRedmineProjectId, issue.Project.Id);
            appointment.SetAppointmentCustomId(Constants.FieldRedmineIssueId, issueId);
            appointment.CreateAppointmentLocation(issueId, issue);
            appointment.SetAppointmentState(AppointmentState.Modified);
            appointment.Save();

            // update last used issues
            var issueRef = this.AllIssues[issueId];
            this.LastUsedIssues.RemoveAll(iref => iref == issueRef);
            this.LastUsedIssues.Insert(0, issueRef);
            if (this.LastUsedIssues.Count > 15)
            {
                this.LastUsedIssues = this.LastUsedIssues.Take(15).ToList();
            }
            Settings.Default.LastUsedIssues = string.Join(";", this.LastUsedIssues.Select(iref => iref.IssueId));
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
                var issue = this._issues[dialog.SelectedIssue.IssueId];
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
                            newEntry.SetAppointmentCustomId(Constants.FieldImportedFromRedmine, 0);
                            newEntry.ReminderSet = false;

                            newEntry.CreateAppointmentLocation(issue.Id, issue);
                            newEntry.Subject = dialog.Description;
                            newEntry.Start = currentDate.Date.AddHours(dialog.StartTime.Hour).AddMinutes(dialog.StartTime.Minute);
                            newEntry.End = currentDate.Date.AddHours(dialog.EndTime.Hour).AddMinutes(dialog.EndTime.Minute);

                            newEntry.UpdateAppointmentFields(null, project.Id, issue.Id, activity.Id, DateTime.Now);
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
        /// Gets the total amount of hours from a list of appointment items. Appointment items that are for the overtime issue are not counted. 
        /// </summary>
        /// <param name="appointmentItems">The appointment items.</param>
        /// <returns>The total amount of hours.</returns>
        private static double GetWorkTime(List<AppointmentItem> appointmentItems)
        {
            var overTimeIssueId = Settings.Default.RedmineUseOvertimeIssue;

            return
                appointmentItems.Where(
                    app =>
                    app.GetAppointmentCustomId(Constants.FieldRedmineIssueId).HasValue
                    && app.GetAppointmentCustomId(Constants.FieldRedmineIssueId) != overTimeIssueId).Sum(app => app.Duration) / 60.0;
        }

        /// <summary>
        /// Creates a time entry for an appointment in redmine
        /// </summary>
        /// <param name="item">
        /// The item to transfer
        /// </param>
        private void CreateInRedmine(AppointmentItem item)
        {
            var updateTime = DateTime.Now;
            var entry = this.CreateTimeEntryFromAppointment(item, updateTime);
            if (entry != null)
            {
                entry.Id = 0;
                var resultEntry = this._redmineManager.CreateObject(entry);

                // update the appointment properties
                item.SetAppointmentCustomId(Constants.FieldRedmineTimeEntryId, resultEntry.Id);
                if (resultEntry.Issue.Id != Settings.Default.RedmineUseOvertimeIssue)
                {
                    item.SetAppointmentState(AppointmentState.Synchronized);
                }
                else
                {
                    item.SetAppointmentState(AppointmentState.SynchronizedOvertime);
                }
                item.SetAppointmentModificationDate(resultEntry.UpdatedOn.Value);
                item.Save();
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
        private TimeEntry CreateTimeEntryFromAppointment(AppointmentItem item, DateTime updateTime)
        {
            var entryId = item.GetAppointmentCustomId(Constants.FieldRedmineTimeEntryId);
            var projectId = item.GetAppointmentCustomId(Constants.FieldRedmineProjectId);
            var issueId = item.GetAppointmentCustomId(Constants.FieldRedmineIssueId);
            var activityId = item.GetAppointmentCustomId(Constants.FieldRedmineActivityId);

            // mandatory fields have to be set
            if (activityId == null)
            {
                // get default activity
                var defaultAct = this.GetDefaultActivity();

                // update appointment
                item.SetAppointmentCustomId(Constants.FieldRedmineActivityId, defaultAct.Id);
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
                    item.SetAppointmentCustomId(Constants.FieldRedmineIssueId, issueGuess.Id);
                    item.Save();

                    issueId = issueGuess.Id;
                    projectId = issueGuess.Project.Id;
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
                item.SetAppointmentCustomId(Constants.FieldRedmineProjectId, issue.Project.Id);
                item.Save();
                projectId = issue.Project.Id;
            }

            // create new time entry
            var timeEntryToCreate = new TimeEntry
                                        {
                                            Activity =
                                                new IdentifiableName
                                                    {
                                                        Id = activityId.Value,
                                                        Name = this._activities[activityId.Value].Name
                                                    },
                                            Issue = new IdentifiableName { Id = issueId.Value, Name = this._issues[issueId.Value].Subject },
                                            Project =
                                                new IdentifiableName { Id = projectId.Value, Name = this._projects[projectId.Value].Name },
                                            User = new IdentifiableName { Id = this.CurrentUser.Id },
                                            Comments = item.Subject,
                                            CreatedOn = DateTime.Now,
                                            UpdatedOn = updateTime,
                                            SpentOn = item.Start.Date,
                                            Hours = 0
                                        };

            // create start and end time custom fields
            var startTimeField = new IssueCustomField { Name = "Start", Id = 1 };
            var startTimeValue = new CustomFieldValue { Info = item.Start.ToString("HH:mm") };
            startTimeField.Values = new List<CustomFieldValue> { startTimeValue };
            timeEntryToCreate.CustomFields = new List<IssueCustomField> { startTimeField };

            var endTimeField = new IssueCustomField { Name = "End", Id = 2 };
            var endTimeValue = new CustomFieldValue { Info = item.End.ToString("HH:mm") };
            endTimeField.Values = new List<CustomFieldValue> { endTimeValue };
            timeEntryToCreate.CustomFields.Add(endTimeField);

            // set the hours
            if (issueId.Value != Settings.Default.RedmineUseOvertimeIssue)
            {
                timeEntryToCreate.Hours = (decimal)(item.End - item.Start).TotalHours;
            }

            // set the Id if there is already one
            if (entryId != null)
            {
                timeEntryToCreate.Id = entryId.Value;
            }

            return timeEntryToCreate;
        }

        /// <summary>
        /// Deletes the time entry that is associated with an <see cref="AppointmentItem"/> from redmine.
        /// </summary>
        /// <param name="item">The item for which to delete the corresponding time entry from redmine.</param>
        private void DeleteTimeEntryInRedmine(AppointmentItem item)
        {
            var timeEntryId = item.GetAppointmentCustomId(Constants.FieldRedmineTimeEntryId);
            if (!timeEntryId.HasValue)
            {
                // if the item has no time entry id, it is not in redmine, thus consider the deletion successful.
                return;
            }

            this._redmineManager.DeleteObject<TimeEntry>(timeEntryId.Value.ToString(), new NameValueCollection());
        }

        /// <summary>
        /// Download the available activities
        /// </summary>
        private void DownloadActivities()
        {
            var parameters = new NameValueCollection { { "limit", "100" } };
            this._activities = this._redmineManager.GetTotalObjectList<TimeEntryActivity>(parameters).ToDictionary(act => act.Id, act => act);
        }

        /// <summary>
        /// Downloads all issues for a specified set of projects.
        /// </summary>
        /// <param name="newProjects">The projects that were not previously known to scorpio, which have to have all their issues downloaded</param>
        /// <returns>A dictionary that contains all issues for the provided projects, identified by their issue id.</returns>
        private Dictionary<int, Issue> DownloadAllIssuesForNewProjects(IEnumerable<Project> newProjects)
        {
            var result = new Dictionary<int, Issue>();
            foreach (var project in newProjects)
            {
                var parameters = new NameValueCollection
                                     {
                                         { "limit", "100" },
                                         { "status_id", "*" },
                                         { "project_id", string.Format("{0}", project.Id) }
                                     };

                var newIssues = this._redmineManager.GetTotalObjectList<Issue>(
                    parameters,
                    (cur, total) =>
                    Globals.ThisAddIn.SyncState.Status =
                    string.Format("Lade Issues für neues Projekt {0} ({1}/{2})", project.Identifier, Math.Min(cur, total), total));

                foreach (var issue in newIssues)
                {
                    if (!result.ContainsKey(issue.Id))
                    {
                        result.Add(issue.Id, issue);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Downloads all redmine issues
        /// </summary>
        /// <param name="newProjects">The projects that are new in this run of downloading issues. 
        /// They will have all Issues loaded, independent of the last change time of the issue.</param>
        private void DownloadIssues(IEnumerable<Project> newProjects)
        {
            // get the issue cache
            var lastSync = Settings.Default.LastIssueSyncDate;
            var knownIssueList = LocalCache.ReadObject(LocalCache.KnownIssues) as List<Issue> ?? new List<Issue>();

            // get new issues since that date minus one
            var parameters = new NameValueCollection { { "limit", "1" }, { "status_id", "*" } };
            if (knownIssueList.Any())
            {
                parameters.Add("updated_on", ">=" + lastSync.Date.AddDays(-2).ToString("yyyy-MM-dd"));
            }

            // parameters.Add("status_id", "open");

            var newIssues =
                this._redmineManager.GetTotalObjectList<Issue>(
                    parameters,
                    (cur, total) => Globals.ThisAddIn.SyncState.Status = string.Format("Lade Issues ({0}/{1})", Math.Min(cur, total), total))
                    .ToDictionary(i => i.Id, i => i);

            // add known issues if not retrieved by download
            foreach (var issue in knownIssueList)
            {
                if (!newIssues.ContainsKey(issue.Id))
                {
                    newIssues.Add(issue.Id, issue);
                }
            }

            // Add Issues from new Projects to the list too, without condition of their last change date (#14644)
            var allIssuesOfNewProjects = this.DownloadAllIssuesForNewProjects(newProjects);
            foreach (var issue in allIssuesOfNewProjects)
            {
                if (!newIssues.ContainsKey(issue.Key))
                {
                    newIssues.Add(issue.Key, issue.Value);
                }
            }

            var success = LocalCache.WriteObject(LocalCache.KnownIssues, newIssues.Values.ToList());
            if (success)
            {
                Settings.Default.LastIssueSyncDate = DateTime.Now.Date;
                Settings.Default.Save();
            }

            this._issues = newIssues;
        }

        /// <summary>
        /// Downloads all redmine projects
        /// </summary>
        /// <returns>The projects that are downloaded from redmine.</returns>
        private IEnumerable<Project> DownloadProjects()
        {
            var parameters = new NameValueCollection();

            this._projects =
                this._redmineManager.GetTotalObjectList<Project>(
                    parameters,
                    (cur, total) => Globals.ThisAddIn.SyncState.Status = string.Format("Lade Projekte ({0}/{1})", Math.Min(cur, total), total))
                    .ToDictionary(p => p.Id, p => p);

            // Get the known projects from localcache
            var knownProjects = LocalCache.ReadObject(LocalCache.KnownProjects) as List<ProjectInfo> ?? new List<ProjectInfo>();

            // Find those projects which are new 
            var newProjects = this._projects.Values.Where(p => knownProjects.All(kp => p.Id != kp.ProjectId));

            // Write all currently known projects
            LocalCache.WriteObject(
                LocalCache.KnownProjects,
                this._projects.Values.Select(p => new ProjectInfo() { ProjectId = p.Id, ProjectName = p.Name }).ToList());

            return newProjects;
        }

        /// <summary>
        /// Gets the default activity for time entries.
        /// </summary>
        /// <returns>The default activity for time entries.</returns>
        private TimeEntryActivity GetDefaultActivity()
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
        private Issue GetIssueForTimeEntry(TimeEntry entry)
        {
            return this._issues.ContainsKey(entry.Issue.Id) ? this._issues[entry.Issue.Id] : null;
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
                this._redmineManager = new RedmineManager(this.RedmineUrl, Settings.Default.RedmineApiKey);

                // Get user information
                this.CurrentUser = this._redmineManager.GetCurrentUser();

                Globals.ThisAddIn.SyncState.Status = "Lade Projekte...";
                this.DownloadActivities();
                var newProjects = this.DownloadProjects();
                Globals.ThisAddIn.SyncState.Status = "Lade Issues...";
                this.DownloadIssues(newProjects);

                // create displayable issue list
                this.AllIssues = this._issues.ToDictionary(
                    i => i.Key,
                    i =>
                    new IssueProjectInfo
                        {
                            IssueId = i.Key,
                            IssueName = i.Value.Subject,
                            ProjectId = i.Value.Project.Id,
                            ProjectName =
                                this._projects.ContainsKey(i.Value.Project.Id) ? this._projects[i.Value.Project.Id].Name : "???",
                            ProjectShortName =
                                this._projects.ContainsKey(i.Value.Project.Id) ? this._projects[i.Value.Project.Id].Identifier : "???",
                        });

                // create last used issues
                var lastUsed = Settings.Default.LastUsedIssues;
                if (!string.IsNullOrWhiteSpace(lastUsed))
                {
                    this.LastUsedIssues.Clear();
                    foreach (var item in lastUsed.Split(';'))
                    {
                        try
                        {
                            var iid = Convert.ToInt32(item);
                            if (this.AllIssues.ContainsKey(iid))
                            {
                                this.LastUsedIssues.Add(this.AllIssues[iid]);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Error("Could not add issue to list of last used issues.", ex);
                        }
                    }
                }

                // create favorite issues
                var favoriteIssues = Settings.Default.FavoriteIssues;
                if (!string.IsNullOrWhiteSpace(favoriteIssues))
                {
                    this.FavoriteIssues.Clear();
                    foreach (var item in favoriteIssues.Split(';'))
                    {
                        try
                        {
                            var iid = Convert.ToInt32(item);
                            if (this.AllIssues.ContainsKey(iid))
                            {
                                this.FavoriteIssues.Add(this.AllIssues[iid]);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Error("Could not add issue to list of last used issues.", ex);
                        }
                    }
                }

                Globals.ThisAddIn.SyncState.Status = "Verbunden";
                this.CanSyncTimeEntries = true;
            }
            catch (Exception ex)
            {
                Log.Error("Could not connect to Redmine.", ex);
                this.CurrentUser = null;
                this.CanSyncTimeEntries = false;
                Globals.ThisAddIn.SyncState.Status = "Nicht verbunden";
            }
            finally
            {
                this.IsConnecting = false;
                Globals.ThisAddIn.SyncState.RaiseConnectionChanged();
            }
        }

        /// <summary>
        /// Is invoked before an appointment item is deleted
        /// </summary>
        /// <param name="item">The item</param>
        /// <param name="cancel">Cancel deletion?</param>
        private void ItemOnBeforeDelete(object item, ref bool cancel)
        {
            var appItem = (AppointmentItem)item;

            var rid = appItem.GetAppointmentCustomId(Constants.FieldRedmineTimeEntryId);
            if (rid != null && rid.Value > 0)
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
            if (name == "Start" || name == "End" || name == "Subject")
            {
                if (!item.IsModifiedSet())
                {
                    item.SetAppointmentState(AppointmentState.Modified);
                    item.Save();
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
             * synchronizer sets another custom field: Constants.FieldImportedFromRedmine. This indicates that we must not 
             * change the state to modified here.
             */
            var importedFromRedmine = appItem.GetAppointmentCustomId(Constants.FieldImportedFromRedmine);
            if (appItem.GetAppointmentCustomId(Constants.FieldRedmineIssueId).HasValue
                && !(importedFromRedmine.HasValue && importedFromRedmine.Value == 1))
            {
                appItem.SetAppointmentState(AppointmentState.Modified);
            }
            else
            {
                appItem.SetAppointmentCustomId(Constants.FieldImportedFromRedmine, null);
            }

            appItem.Save();
            this.RegisterAppointment(appItem);
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
            var parameters = new NameValueCollection
                                 {
                                     { "spent_on", string.Format("><{0:yyyy-MM-dd}|{1:yyyy-MM-dd}", startDate, endDate) },
                                     { "user_id", "me" },
                                     { "limit", "100" }
                                 };

            var retrievedItems = new List<TimeEntry>();

            try
            {
                var items = this._redmineManager.GetTotalObjectList<TimeEntry>(
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
                    var redmineId = appointment.GetAppointmentCustomId(Constants.FieldRedmineTimeEntryId);
                    if (redmineId == null || appointment.IsAppointmentCopied())
                    {
                        // delete items that does not exist in redmine
                        appointment.SetAppointmentCustomId(Constants.FieldRedmineTimeEntryId, null);
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
                                appointment.SetAppointmentModificationDate(redmineEntry.UpdatedOn.Value);
                            }

                            if (redmineEntry.Issue.Id != Settings.Default.RedmineUseOvertimeIssue)
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
                            appointment.SetAppointmentCustomId(Constants.FieldRedmineTimeEntryId, null);
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
                    newEntry.SetAppointmentCustomId(Constants.FieldImportedFromRedmine, 1);
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
                    Globals.ThisAddIn.SyncState.RaiseConnectionChanged();

                    // If the item was conflicted, try to reset its previous state, 
                    // modified or deleted, s.t. we can do another synchronization.
                    if (item.IsSyncErrorSet())
                    {
                        item.ResetPreviousState();
                    }

                    this.SaveTimeEntry(item);
                }
                catch (Exception ex)
                {
                    // put the exception details to the appointment body
                    item.AppendToBody(ex.ToString());
                    item.SetAppointmentState(AppointmentState.SyncError);
                    item.Save();

                    var message = string.Format(
                        "Zeiteintrag {0} - {1}, {2}: {3} konnte nicht synchronisiert werden.",
                        item.Start,
                        item.End,
                        item.Location,
                        item.Subject);
                    errors.Add(message);
                    Log.Error(message, ex);
                }
            }

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
            if (item.IsModifiedSet())
            {
                var timeEntryId = item.GetAppointmentCustomId(Constants.FieldRedmineTimeEntryId);
                if (timeEntryId.HasValue && !item.IsAppointmentCopied())
                {
                    this.UpdateInRedmine(item);
                }
                else
                {
                    this.CreateInRedmine(item);
                    item.MarkAsNotCopied();
                }
            }
            else if (item.IsDeletedSet())
            {
                this.DeleteTimeEntryInRedmine(item);

                item.SetAppointmentCustomId(Constants.FieldRedmineTimeEntryId, null);
                item.Delete();
            }
            else
            {
                Log.Warn(string.Format("Item is assumed modified, but has no category set for modified or deleted: {0}", item.Location));
            }
        }

        /// <summary>
        /// Tries to guess the correct issue from the appointment subject line
        /// </summary>
        /// <param name="item">The appointment item</param>
        /// <returns>The issue or null if no issue could be found</returns>
        private Issue TryGuessIssue(AppointmentItem item)
        {
            if (RxIssueNumber.IsMatch(item.Subject))
            {
                try
                {
                    var m = RxIssueNumber.Match(item.Subject);
                    var issueIdStr = m.Groups["IssueNumber"].Captures[0].Value;
                    var id = Convert.ToInt32(issueIdStr);
                    if (this._issues.ContainsKey(id))
                    {
                        return this._issues[id];
                    }
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
        /// Updates the hours status for the current view
        /// </summary>
        /// <param name="from">Start date</param>
        /// <param name="to">End date</param>
        private void UpdateHoursInView(DateTime from, DateTime to)
        {
            // This is ugly, because it will frequently get updated by a timer, regardless of whether something has actually changed. 
            // This will be changed in a later version, by employing change events that are fired whenever a timeentry is created/updated/etc.
            var now = DateTime.Now;

            var appointmentsInView = this.Calendar.GetAppointmentsInRange(from, to.Date.AddDays(1).AddSeconds(-1));
            var hoursInView = GetWorkTime(appointmentsInView);
            Globals.ThisAddIn.SyncState.HoursInView = hoursInView;

            var appointmentsInDay = this.Calendar.GetAppointmentsInRange(DateTimeHelper.StartOfDay(now), DateTimeHelper.EndOfDay(now));
            var hoursInDay = GetWorkTime(appointmentsInDay);
            Globals.ThisAddIn.SyncState.HoursInDay = hoursInDay;

            var appointmentsInWeek = this.Calendar.GetAppointmentsInRange(DateTimeHelper.StartOfWeek(now), DateTimeHelper.EndOfWeek(now));
            var hoursInWeek = GetWorkTime(appointmentsInWeek);
            Globals.ThisAddIn.SyncState.HoursInWeek = hoursInWeek;

            var appointmentsInMonth = this.Calendar.GetAppointmentsInRange(DateTimeHelper.StartOfMonth(now), DateTimeHelper.EndOfMonth(now));
            var hoursInMonth = GetWorkTime(appointmentsInMonth);
            Globals.ThisAddIn.SyncState.HoursInMonth = hoursInMonth;
        }

        /// <summary>
        /// Updates a time entry in redmine from an appointment
        /// </summary>
        /// <param name="item">
        /// The item to transfer
        /// </param>
        private void UpdateInRedmine(AppointmentItem item)
        {
            var updateTime = DateTime.Now;
            var entry = this.CreateTimeEntryFromAppointment(item, updateTime);
            if (entry != null)
            {
                // send the update to redmine
                this._redmineManager.DeleteObject<TimeEntry>(entry.Id.ToString(), new NameValueCollection());

                entry.Id = 0;
                var resultEntry = this._redmineManager.CreateObject(entry);

                item.SetAppointmentCustomId(Constants.FieldRedmineTimeEntryId, resultEntry.Id);
                if (resultEntry.Issue.Id != Settings.Default.RedmineUseOvertimeIssue)
                {
                    item.SetAppointmentState(AppointmentState.Synchronized);
                }
                else
                {
                    item.SetAppointmentState(AppointmentState.SynchronizedOvertime);
                }

                item.SetAppointmentModificationDate(resultEntry.UpdatedOn.Value);
                item.Save();
            }
        }

        /// <summary>
        /// Method which is called periodically by the <see cref="_updateTimer"/>. It keeps the display of 
        /// logged hours in the ribbon bar synchronized with the calendar view.
        /// </summary>
        /// <param name="sender">The timer</param>
        /// <param name="args">The elapsed event args</param>
        private void UpdateTimerCallback(object sender, ElapsedEventArgs args)
        {
            /*
             * TODO: We should try to get rid of the update timer. However, Outlook does not seem 
             * to provide a callback for change of the displayed dates in the calendar view:
             * http://stackoverflow.com/questions/32693475/outlook-2013-vsto-get-calendar-selected-range-callback
             * 
             * TODO: This update method might go out of the redmine synchronizer
             */

            var dates = Globals.ThisAddIn.CalendarState.GetDisplayDates();

            if (dates == null || dates.Length == 0)
            {
                return;
            }
            this.UpdateHoursInView(dates.Min(), dates.Max());
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Gets the user which is currently logged into redmine.
        /// </summary>
        public User CurrentUser { get; private set; }

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
                return this._redmineManager != null && this.CurrentUser != null;
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
        public IDictionary<int, IssueProjectInfo> AllIssues { get; private set; }

        /// <summary>
        /// Gets a list of previously used issues
        /// </summary>
        public List<IssueProjectInfo> LastUsedIssues { get; private set; }

        /// <summary>
        /// Gets a list of issues that were marked as favorite issues.
        /// </summary>
        public List<IssueProjectInfo> FavoriteIssues { get; private set; }

        /// <summary>
        /// Saves the favorite issues to the settings.
        /// </summary>
        /// <param name="issues">The issue information of the issues which shall become the new favorite issues.</param>
        public void UpdateFavoriteIssues(List<IssueProjectInfo> issues)
        {
            this.FavoriteIssues = issues;

            Settings.Default.FavoriteIssues = string.Join(";", this.FavoriteIssues.Select(iref => iref.IssueId));
            Settings.Default.Save();
        }

        /// <summary>
        /// Gets the URL to the redmine system
        /// </summary>
        public string RedmineUrl
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
                return this.CurrentUser != null ? string.Format("{0} {1}", this.CurrentUser.FirstName, this.CurrentUser.LastName) : "";
            }
        }

        #endregion
    }
}