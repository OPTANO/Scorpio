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

namespace Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Linq;
 
    using global::Redmine.Net.Api;
    using global::Redmine.Net.Api.Types;

    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Properties;
    using Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource.Exceptions;

    /// <summary>
    /// A concrete redmine manager instance using the remine manager api for syncing with redmine
    /// </summary>
    public class RedmineManagerInstance : IExternalSource
    {
        #region Fields

        /// <summary>
        /// The redmine manager
        /// </summary>
        private readonly RedmineManager _redmineApi;

        /// <summary>
        /// Backing field for fast access to the user info, the element only has to be queried once.
        /// The manager is reset, if the connection changes, hence this is ok
        /// </summary>
        private UserInfo _userInfo;
        
        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="RedmineManagerInstance"/> class.
        /// </summary>
        /// <param name="address">
        /// The host address.
        /// </param>
        /// <param name="apiKey">
        /// The api key.
        /// </param>
        /// <param name="limit">the limit to use</param>
        public RedmineManagerInstance(string address, string apiKey, int limit = 100)
        {
            this._redmineApi = new RedmineManager(address, apiKey);
            this._redmineApi.PageSize = 50;
            this.Limit = limit;
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Gets or sets the limit for a request made
        /// </summary>
        public int Limit { get; set; }

        /// <summary>
        /// Method to create an object from the given time entry
        /// </summary>
        /// <param name="item">the time entry</param>
        /// <returns>the object created</returns>
        public TimeEntryInfo CreateObject(TimeEntryInfo item)
        {
            try
            {
                // create the time entry to send to redmine
                var timeEntry = this.CreateTimeEntryFromInfoObject(item);

                // write the object to redmine
                var createdObject = this._redmineApi.CreateObject(timeEntry);

                // update object and return it
                item.Id = createdObject.Id;
                return item;
            }
            catch (RedmineException redmineException)
            {
                var typeNumber = redmineException.HResult;
                if (redmineException.Message.Contains("Die zugrunde liegende Verbindung wurde geschlossen"))
                {
                    throw new ConnectionException(redmineException) { IdentifierNumber = typeNumber };
                }
                else
                {
                    throw new CrudException(OperationType.Create, item, redmineException) { IdentifierNumber = typeNumber };
                }
            }
        }

        /// <summary>
        /// Method to create the a time entry from the info object
        /// </summary>
        /// <param name="item">the time entry info to update the infos from</param>
        /// <returns>the time entry which can be sent to redmine</returns>
        internal TimeEntry CreateTimeEntryFromInfoObject(TimeEntryInfo item)
        {
            // get base info
            var user = this.GetCurrentUser();
            var activityInfo = item.ActivityInfo;
            var projectInfo = item.ProjectInfo;
            var issueInfo = item.IssueInfo;

            // create new time entry for redmine
            var timeEntry = new TimeEntry
                                {
                                    Activity = new IdentifiableName { Id = activityInfo.Id, Name = activityInfo.Name },
                                    Issue = new IdentifiableName { Id = issueInfo.Id, Name = issueInfo.Name },
                                    Project = new IdentifiableName { Id = projectInfo.Id, Name = projectInfo.Name },
                                    User = new IdentifiableName { Id = user.Id },
                                    Comments = item.Name,
                                    CreatedOn = DateTime.Now,
                                    UpdatedOn = item.UpdateTime,
                                    SpentOn = item.StartDateTime.Date,
                                    Hours = 0
                                };

            // create start and end time custom fields
            var startTimeField = new IssueCustomField { Name = "Start", Id = 1 };
            var startTimeValue = new CustomFieldValue { Info = item.StartDateTime.ToString("HH:mm") };
            startTimeField.Values = new List<CustomFieldValue> { startTimeValue };
            timeEntry.CustomFields = new List<IssueCustomField> { startTimeField };

            var endTimeField = new IssueCustomField { Name = "End", Id = 2 };
            var endTimeValue = new CustomFieldValue { Info = item.EndDateTime.ToString("HH:mm") };
            endTimeField.Values = new List<CustomFieldValue> { endTimeValue };
            timeEntry.CustomFields.Add(endTimeField);

            // set the hours
            if (issueInfo.Id != Settings.Default.RedmineUseOvertimeIssue)
            {
                timeEntry.Hours = (decimal)(item.EndDateTime - item.StartDateTime).TotalHours;
            }

            // set the Id if there is already one
            if (item.Id > 0)
            {
                timeEntry.Id = item.Id;
            }
            return timeEntry;
        }

        /// <summary>
        /// Method to update the given object with the data contained in the object. The object has the id set, all other values should be updated
        /// </summary>
        /// <param name="item">the entry to update</param>
        /// <returns>the time entry info of the updated object</returns>
        public TimeEntryInfo UpdateObject(TimeEntryInfo item)
        {
            try
            {
                // create the time entry to send to redmine
                var timeEntry = this.CreateTimeEntryFromInfoObject(item);

                // write the object to redmine
                this._redmineApi.UpdateObject(timeEntry.Id.ToString(), timeEntry);

                // update object and return it
                return item;
            }
            catch (RedmineException redmineException)
            {
                var typeNumber = redmineException.HResult;
                if (redmineException.Message.Contains("Die zugrunde liegende Verbindung wurde geschlossen"))
                {
                    throw new ConnectionException(redmineException) { IdentifierNumber = typeNumber };
                }
                else
                {
                    throw new CrudException(OperationType.Update, item, redmineException) { IdentifierNumber = typeNumber };
                }
            }
        }

        /// <summary>
        /// Method to delete an object by its name
        /// </summary>
        /// <param name="objectIdentifier">the identifier of the object</param>
        /// <param name="parameters">the name value collection</param>
        public void DeleteTimeEntry(int? objectIdentifier, DataSourceParameter parameters)
        {
            if (objectIdentifier.HasValue)
            {
                try
                {
                    this._redmineApi.DeleteObject<TimeEntry>(objectIdentifier.ToString(), this.GetParametersForQuery(parameters));
                }
                catch (RedmineException redmineException)
                {
                    if (redmineException.Message.Equals("Not Found"))
                    {
                        // nothing to do, the item does no longer exist in redmine, so we can delete it in the application
                    }
                    else
                    {
                        var typeNumber = redmineException.HResult;
                        if (redmineException.Message.Contains("Die zugrunde liegende Verbindung wurde geschlossen"))
                        {
                            throw new ConnectionException(redmineException) { IdentifierNumber = typeNumber };
                        }
                        else
                        {
                            throw new CrudException(OperationType.Delete, null, redmineException) { IdentifierNumber = typeNumber };
                        }
                    }
                }
            }
        }
        
        /// <summary>
        /// Method to get the current user logged in to redmine
        /// </summary>
        /// <returns>the user</returns>
        public UserInfo GetCurrentUser()
        {
            if (this._userInfo == null)
            {
                var redmineUser = this._redmineApi.GetCurrentUser();
                var name = string.Format("{0} {1} ({2})", redmineUser.FirstName, redmineUser.LastName, redmineUser.Login);
                var user = new UserInfo() { Id = redmineUser.Id, Name = name };
                this._userInfo = user;
            }
            return this._userInfo;
        }

        /// <summary>
        /// Method to get the total object list of all projects
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        public IList<ProjectInfo> GetTotalProjectList(DataSourceParameter parameters, Action<int, int> statusCallback = null)
        {
            var projects = this._redmineApi.GetTotalObjectList<Project>(this.GetParametersForQuery(parameters), statusCallback);
            var projectInfos = projects.Select(p => new ProjectInfo() { Id = p.Id, Name = p.Identifier }).ToList();

            return projectInfos;
        }

        /// <summary>
        /// Method to get the total object list of all issues
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        public IList<IssueInfo> GetTotalIssueInfoList(DataSourceParameter parameters, Action<int, int> statusCallback = null)
        {
            // assert that there is no limit parameter set
            if (parameters.Limit.HasValue)
            {
                throw new ArgumentException("No limit parameter may be set for this method.");
            }

            // get parameter and issues
            var parameter = this.GetParametersForQuery(parameters);
            var issues = this._redmineApi.GetTotalObjectList<Issue>(parameter, statusCallback);

            // convert issues
            var issueInfos =
                issues.Select(i => new IssueInfo() { Id = i.Id, Name = i.Subject, ProjectShortName = i.Project.Name, ProjectId = i.Project.Id, })
                    .ToList();
            
            return issueInfos;
        }

        /// <summary>
        /// Method to get the list of all issues matching the given parameter
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <returns>the list containing all objects</returns>
        public IList<IssueInfo> GetIssueInfoList(DataSourceParameter parameters)
        {
            var parameter = this.GetParametersForQuery(parameters);
            var issues = this._redmineApi.GetObjectList<Issue>(parameter);

            var issueInfos =
                issues.Select(i => new IssueInfo() { Id = i.Id, Name = i.Subject, ProjectShortName = i.Project.Name, ProjectId = i.Project.Id, })
                    .ToList();

            return issueInfos;
        }

        /// <summary>
        /// Method to get the total object list of all activities
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        public IList<ActivityInfo> GetTotalActivityInfoList(DataSourceParameter parameters, Action<int, int> statusCallback = null)
        {
            var activities = this._redmineApi.GetTotalObjectList<TimeEntryActivity>(this.GetParametersForQuery(parameters), statusCallback);
            var activityInfos = activities.Select(a => new ActivityInfo() { Id = a.Id, Name = a.Name, IsDefault = a.IsDefault }).ToList();
            return activityInfos;
        }

        /// <summary>
        /// Method to get the total object list of all time entries
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        public IList<TimeEntryInfo> GetTotalTimeEntryInfoList(DataSourceParameter parameters, Action<int, int> statusCallback = null)
        {
            var timeEntries = this._redmineApi.GetTotalObjectList<TimeEntry>(this.GetParametersForQuery(parameters), statusCallback);
            var timeEntryInfos = timeEntries.Select(
                te =>
                    {
                        var timeOfAction = te.SpentOn.Value.Date;
                        var customFields = te.CustomFields;
                        var startField = customFields.FirstOrDefault(f => f.Name == "Start");
                        if (startField != null)
                        {
                            var startText = startField.Values.First().Info;
                            var startTextSplit = startText.Split(':');
                            if (startTextSplit.Length == 2)
                            {
                                var minutes = 0;
                                var hours = 0;
                                var parseSuccessful = int.TryParse(startTextSplit[1], out minutes) && int.TryParse(startTextSplit[0], out hours);

                                if (parseSuccessful)
                                {
                                    var startDateTime = timeOfAction.AddHours(hours).AddMinutes(minutes);

                                    var info = new TimeEntryInfo()
                                                   {
                                                       Id = te.Id,
                                                       Name = te.Comments,
                                                       UpdateTime = te.UpdatedOn.Value,
                                                       StartDateTime = startDateTime,
                                                       EndDateTime = startDateTime.AddHours((double)te.Hours),
                                                       IssueInfo =
                                                           new IssueInfo()
                                                               {
                                                                   Id = te.Issue.Id,
                                                                   Name = te.Issue.Name,
                                                                   ProjectId = te.Project.Id,
                                                                   ProjectShortName = te.Project.Name
                                                               },
                                                       ActivityInfo = new ActivityInfo() { Id = te.Activity.Id, Name = te.Activity.Name },
                                                       ProjectInfo = new ProjectInfo() { Id = te.Project.Id, Name = te.Project.Name, },
                                                   };
                                    return info;
                                }
                            }
                        }
                        return null;
                    }).Where(i => i != null).ToList();
            return timeEntryInfos;
        }
        
        #endregion

        /// <summary>
        /// Method to get the parameter to use for the query
        /// </summary>
        /// <param name="parameter">the parameter given</param>
        /// <returns>the parameter to use</returns>
        internal NameValueCollection GetParametersForQuery(DataSourceParameter parameter)
        {
            var queryParameter = new NameValueCollection();

            if (parameter != null)
            {
                if (parameter.Limit.HasValue)
                {
                    queryParameter.Add("limit", parameter.Limit.ToString());
                }
                else if (parameter.UseLimit.HasValue && parameter.UseLimit.Value)
                {
                    queryParameter.Add("limit", this.Limit.ToString());
                }

                if (parameter.StatusId.HasValue)
                {
                    queryParameter.Add("status_id", parameter.StatusId.Value == -1 ? "*" : parameter.StatusId.Value.ToString());
                }
                if (parameter.ProjectId.HasValue)
                {
                    queryParameter.Add("project_id", parameter.ProjectId.Value.ToString());
                }
                if (parameter.UpdateStartDateTime.HasValue)
                {
                    queryParameter.Add("updated_on", ">=" + parameter.UpdateStartDateTime.Value.ToString("yyyy-MM-dd"));
                }
                if (parameter.UserId.HasValue)
                {
                    queryParameter.Add("user_id", parameter.UserId.Value == -1 ? "me" : parameter.UserId.Value.ToString());
                }
                if (parameter.SpentDateTimeTuple != null)
                {
                    var tuple = parameter.SpentDateTimeTuple;
                    queryParameter.Add("spent_on", string.Format("><{0:yyyy-MM-dd}|{1:yyyy-MM-dd}", tuple.Item1, tuple.Item2));
                }
                if (parameter.IssueId.HasValue)
                {
                    queryParameter.Add("issue_id", parameter.IssueId.Value.ToString());
                }
            }

            return queryParameter;
        }
    }
}