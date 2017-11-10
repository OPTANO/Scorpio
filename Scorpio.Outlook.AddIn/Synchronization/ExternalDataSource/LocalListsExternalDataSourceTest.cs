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
    
    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource.Exceptions;

    /// <summary>
    /// An implementation of the external data source class working against local test data, can be used for tests to simulate special events
    /// </summary>
    public class LocalListsExternalDataSourceTest : IExternalSource
    {
        #region Fields

        /// <summary>
        /// The issue list
        /// </summary>
        private readonly List<IssueInfo> _issues;

        /// <summary>
        /// The projects list
        /// </summary>
        private readonly List<ProjectInfo> _projects;

        /// <summary>
        /// The time entries list
        /// </summary>
        private readonly List<TimeEntryInfo> _timeEntries;

        /// <summary>
        /// The time entry action list
        /// </summary>
        private readonly List<ActivityInfo> _timeEntryActivities;

        /// <summary>
        /// The user
        /// </summary>
        private readonly UserInfo _user;

        /// <summary>
        /// The issue counter
        /// </summary>
        private int _timeEntryCounter;

        #endregion

        /// <summary>
        /// The number of projects to be created 
        /// </summary>
        private int _numberProjects = 20;

        /// <summary>
        /// The number of issues to be created per project
        /// </summary>
        private int _numberIssuesPerProject = 1500;

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="LocalListsExternalDataSourceTest"/> class.
        /// </summary>
        public LocalListsExternalDataSourceTest()
        {
            // initialize lists 
            this._issues = new List<IssueInfo>();
            this._projects = new List<ProjectInfo>();
            this._timeEntries = new List<TimeEntryInfo>();
            this._timeEntryActivities = new List<ActivityInfo>();
            this.ProjectsWithWatchedIssueStatus = new HashSet<int>();
            
            // add activities
            var defaultActivity = new ActivityInfo() { Id = 1, Name = "Standard", IsDefault = true };
            var otherActivity = new ActivityInfo() { Id = 2, Name = "was anderes", IsDefault = true };
            this._timeEntryActivities.Add(defaultActivity);
            this._timeEntryActivities.Add(otherActivity);

            // add user
            this._user = new UserInfo() { Id = 42, Name = "Scorpio Testuser", };

            // initialize counter
            var issueId = 0;
            this._timeEntryCounter = 0;

            // add projects and issues for them
            for (var projectCounter = 0; projectCounter < this._numberProjects; projectCounter++)
            {
                var project = new ProjectInfo()
                                  {
                                      Name = string.Format("{0} TestProject", projectCounter),
                                      Id = projectCounter
                                  };
                this._projects.Add(project);
                
                for (var issueCounter = 0; issueCounter < this._numberIssuesPerProject; issueCounter++)
                {
                    this._issues.Add(
                        new IssueInfo()
                            {
                                ProjectId = project.Id.Value,
                                ProjectShortName = project.Name,
                                Id = issueId,
                                Name = string.Format("{0} Issue", issueId)
                            });
                    issueId += 1;
                }
            }
            
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Gets or sets the limit for a request made
        /// </summary>
        public int Limit { get; set; }

        /// <inheritdoc />
        public HashSet<int> ProjectsWithWatchedIssueStatus { get; }

        /// <summary>
        /// Method to create an object from teh given time entry
        /// </summary>
        /// <param name="entry">the time entry</param>
        /// <returns>the object created</returns>
        public TimeEntryInfo CreateObject(TimeEntryInfo entry)
        {
            // for test purposes, to see what happens
            // throw new ConnectionException(null);
            // throw new CrudException(OperationType.Create, entry, null);
            // throw new AccessViolationException();
            entry.Id = this._timeEntryCounter++;
            this._timeEntries.Add(entry);
            return entry;
        }

        /// <summary>
        /// Method to update the given object with the data contained in the object. The object has the id set, all other values should be updated
        /// </summary>
        /// <param name="entry">the entry to update</param>
        /// <returns>the time entry info of the updated object</returns>
        public TimeEntryInfo UpdateObject(TimeEntryInfo entry)
        {
            this._timeEntries.RemoveAll(e => object.Equals(e.Id, entry.Id));
            this._timeEntries.Add(entry);
            return entry;
        }

        /// <summary>
        /// Method to delete an object by its name
        /// </summary>
        /// <param name="objectIdentifier">the identifier of the object</param>
        /// <param name="nameValueCollection">the name value collection</param>
        public void DeleteTimeEntry(int? objectIdentifier, DataSourceParameter nameValueCollection)
        {
            if (objectIdentifier != null)
            {
                this._timeEntries.RemoveAll(new Predicate<TimeEntryInfo>(i => i.Id == objectIdentifier));
            }
        }

        /// <summary>
        /// Method to get the current user logged in to redmine
        /// </summary>
        /// <returns>the user</returns>
        public UserInfo GetCurrentUser()
        {
            return this._user;
        }

        /// <summary>
        /// Method to get the total object list of all projects
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        public IList<ProjectInfo> GetTotalProjectList(DataSourceParameter parameters, Action<int, int> statusCallback = null)
        {
            return this._projects;
        }

        /// <summary>
        /// Method to get the total object list of all issues
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        public IList<IssueInfo> GetTotalIssueInfoList(DataSourceParameter parameters, Action<int, int> statusCallback = null)
        {
            return this._issues;
        }

        /// <summary>
        /// Method to get the list of all issues matching the given parameter
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <returns>the list containing all objects</returns>
        public IList<IssueInfo> GetIssueInfoList(DataSourceParameter parameters)
        {
            return this._issues;
        }

        /// <summary>
        /// Method to get the total object list of all activities
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        public IList<ActivityInfo> GetTotalActivityInfoList(DataSourceParameter parameters, Action<int, int> statusCallback = null)
        {
            return this._timeEntryActivities;
        }

        /// <summary>
        /// Method to get the total object list of all time entries
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        public IList<TimeEntryInfo> GetTotalTimeEntryInfoList(DataSourceParameter parameters, Action<int, int> statusCallback = null)
        {
            return this._timeEntries;
        }
        
        #endregion
    }
}