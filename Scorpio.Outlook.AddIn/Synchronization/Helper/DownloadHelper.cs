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

namespace Scorpio.Outlook.AddIn.Synchronization.Helper
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using DevExpress.Mvvm.Native;

    using log4net;

    using Scorpio.Outlook.AddIn.Cache;
    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Properties;
    using Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource;

    /// <summary>
    /// Helper class for downloading and updating the issue list using an external data source
    /// </summary>
    public static class DownloadHelper
    {
        #region Static Fields

        /// <summary>
        /// The logger.
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(DownloadHelper));

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Method to get the list of all activity infos
        /// </summary>
        /// <param name="source">the external data source</param>
        /// <returns>the activities</returns>
        public static Dictionary<int, ActivityInfo> GetActivityInfos(IExternalSource source)
        {
            var parameters = new DataSourceParameter() { UseLimit = false };
            var activities = source.GetTotalActivityInfoList(parameters).ToDictionary(act => act.Id.Value, act => act);
            return activities;
        }

        /// <summary>
        /// Downloads all redmine projects and update the cache with the new data
        /// </summary>
        /// <param name="source">the data source to use</param>
        /// <returns>The projects that are downloaded from redmine.</returns>
        public static ProjectLists GetCurrentProjectsAndUpdateCache(IExternalSource source)
        {
            if (Globals.ThisAddIn != null)
            {
                Globals.ThisAddIn.SyncState.Status = "Lade Projekte...";
            }

            var parameters = new DataSourceParameter();
            var projects = source.GetTotalProjectList(
                parameters,
                (cur, total) =>
                {
                    if (Globals.ThisAddIn != null)
                    {
                        Globals.ThisAddIn.SyncState.Status = string.Format("Lade Projekte ({0}/{1})", Math.Min(cur, total), total);
                    }
                }).Where(p => p.Id.HasValue).ToDictionary(p => p.Id.Value, p => p);

            // Get the known projects from localcache
            var knownProjects = LocalCache.ReadObject(LocalCache.KnownProjects, new List<ProjectInfo>()) as List<ProjectInfo> ?? new List<ProjectInfo>();

            // Find those projects which are new 
            var newProjects = projects.Values.Except(knownProjects).ToList();

            // Write all currently known projects
            LocalCache.WriteObject(LocalCache.KnownProjects, projects.Values.ToList());

            return new ProjectLists() { AllProjects = projects, NewProjects = newProjects };
        }


        /// <summary>
        /// Method to download the issue list from the given source, known issues are read, updated are done, updated list is returned. In case of an exception, an empty list is returned.
        /// </summary>
        /// <param name="newProjects">the list of new projects to be considered</param>
        /// <param name="currentIssues">the current issue list, which should be updated</param>
        /// <param name="source">the external data source</param>
        /// <returns>the updated list of issues</returns>
        public static Dictionary<int, IssueInfo> DownloadIssues(
            List<ProjectInfo> newProjects,
            IDictionary<int, IssueInfo> currentIssues,
            IExternalSource source)
        {
            if (Globals.ThisAddIn != null)
            {
                Globals.ThisAddIn.SyncState.Status = "Lade Issues...";
            }

            var oldIssues = currentIssues.ToDictionary(k => k.Key, v => v.Value);
            var issues = new Dictionary<int, IssueInfo>();
            try
            {
                // get the issue cache and store the issues as the list of old issues
                var knownIssueList = LocalCache.GetKnownIssueListFromCache();
                if (knownIssueList.Any())
                {
                    oldIssues = knownIssueList.Distinct().Where(i => i.Id.HasValue).ToDictionary(k => k.Id.Value, v => v);
                }

                // stores the current issue information in the issue dictionary
                UpdateCurrentIssues(newProjects, source, knownIssueList, issues);

                // update the issue cache
                LocalCache.UpdateKnownIssuesListInCache(issues);
            }
            catch (Exception exception)
            {
                // if an exception occurs, roll back to the last known issue state
                Log.Error("Fehler während Download Issues", exception);
                issues = oldIssues;
                LocalCache.UpdateKnownIssuesListInCache(issues);
            }

            return issues;
        }

        /// <summary>
        /// The extended reload issue info method, reloads all issues, because reloading for a given issue number where the issue status is done, is not working with the api.
        /// </summary>
        /// <param name="issueId">
        /// The issue id to obtain.
        /// </param>
        /// <param name="source">
        /// The source to use.
        /// </param>
        /// <param name="knownIssueList">
        /// The known issue list.
        /// </param>
        /// <returns>
        /// The <see cref="Dictionary"/> containing all issues.
        /// </returns>
        public static Dictionary<int, IssueInfo> ReloadIssueInfoExtended(int issueId, IExternalSource source, List<IssueInfo> knownIssueList)
        {
            var issues = new Dictionary<int, IssueInfo>();

            // stores the current issue information in the issue dictionary
            UpdateCurrentIssues(new List<ProjectInfo>(), source, knownIssueList, issues);

            // update the issue cache
            LocalCache.UpdateKnownIssuesListInCache(issues);

            return issues;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Downloads all issues for a specified set of projects.
        /// </summary>
        /// <param name="newProjects">The projects that were not previously known to scorpio, which have to have all their issues downloaded</param>
        /// <param name="source">the data source to use</param>
        /// <returns>A dictionary that contains all issues for the provided projects, identified by their issue id.</returns>
        private static Dictionary<int, IssueInfo> DownloadAllIssuesForNewProjects(List<ProjectInfo> newProjects, IExternalSource source)
        {
            // initialize the list of new project
            var allIssuesForNewProjects = new HashSet<IssueInfo>();

            // download the issues for each project
            foreach (var project in newProjects)
            {
                var issuesForProject = DownloadIssuesForProject(source, project);
                issuesForProject.ToList().ForEach(i => allIssuesForNewProjects.Add(i));
                if (issuesForProject.Any(i => !i.Id.HasValue))
                {
                    Log.Warn(string.Format("Project {0} (ID {1}) has at least one issue without id set", project.Name, project.Id));
                }
            }

            // convert the hash set to a dictionary
            var resultDictionary = new Dictionary<int, IssueInfo>();
            allIssuesForNewProjects.Where(i => i.Id.HasValue).ForEach(i => resultDictionary[i.Id.Value] = i);
            
            return resultDictionary;
        }

        /// <summary>
        /// Method to download the issues for the given project
        /// </summary>
        /// <param name="source">the external data source to use</param>
        /// <param name="project">the project to download the issues for</param>
        /// <returns>the list of issues corresponding to the given project</returns>
        private static IList<IssueInfo> DownloadIssuesForProject(IExternalSource source, ProjectInfo project)
        {
            // get parameter and action
            var parameters = new DataSourceParameter() { ProjectId = project.Id, StatusId = -1 };
            Action<int, int> action =
                (i, overall) =>
                    {
                        if (Globals.ThisAddIn == null)
                        {
                            return;
                        }
                        Globals.ThisAddIn.SyncState.Status = string.Format(
                            "Lade Issues für neues Projekt {0} ({1}/{2})",
                            project.Name,
                            Math.Min(i, overall),
                            overall);
                    };
            try
            {
                // load data and add the issues. Note that we load all issues here
                var newIssuesOfProject = source.GetTotalIssueInfoList(parameters, action);
                return newIssuesOfProject;
            }
            catch (Exception exception)
            {
                Log.Error(string.Format("Error while loading project {0}", project.Name), exception);
                return new List<IssueInfo>();
            }
        }

        /// <summary>
        /// Method to get the parameter to use for the download of the issues
        /// </summary>
        /// <param name="issuesKnown">flag indicating if up to now issues are known already</param>
        /// <returns>the parameter to use for the query</returns>
        private static DataSourceParameter GetDataSourceParameterForDownloadOfIssues(bool issuesKnown)
        {
            // get new issues since that date minus one
            var lastSync = Settings.Default.LastIssueSyncDate;
            var parameters = new DataSourceParameter() { UseLimit = true, StatusId = -1 };
            if (issuesKnown)
            {
                Log.Info(string.Format("Last Sync Date set to {0:d}", lastSync.Date.AddDays(-2)));
                parameters.UpdateStartDateTime = lastSync.Date.AddDays(-2);
            }
            return parameters;
        }

        /// <summary>
        /// Method to get the up to date issue information and store them in the issue dictionary
        /// </summary>
        /// <param name="newProjects">the list of new projects to be loaded</param>
        /// <param name="source">the external data source</param>
        /// <param name="knownIssueList">the list of known issues</param>
        /// <param name="issues">the dictionary to store the current state in</param>
        private static void UpdateCurrentIssues(
            List<ProjectInfo> newProjects,
            IExternalSource source,
            List<IssueInfo> knownIssueList,
            Dictionary<int, IssueInfo> issues)
        {
            // initialize the list of all issues
            var allIssues = new List<IssueInfo>();

            // get download parameter
            var issuesKnown = knownIssueList.Any();
            var parameters = GetDataSourceParameterForDownloadOfIssues(issuesKnown);
            if(Globals.ThisAddIn != null)
            {
                Globals.ThisAddIn.SyncState.Status = "Lade neue Issues";
            }

            // get and add new issues
            var newIssueInfos = source.GetIssueInfoList(parameters).Distinct().ToList();
            allIssues.AddRange(newIssueInfos);

            // get and add issues of new projects
            // do not check condition of their last change date (#14644)
            var allIssuesOfNewProjects = DownloadAllIssuesForNewProjects(newProjects, source);
            allIssues.AddRange(allIssuesOfNewProjects.Values);

            // update and add known issues (we only want to add the issues not contained in the new issue list or new project issue list, which can contain an update
            knownIssueList = knownIssueList.Except(newIssueInfos).ToList();
            knownIssueList = knownIssueList.Except(allIssuesOfNewProjects.Values).ToList();
            allIssues.AddRange(knownIssueList);

            // store the values in the local variable
            // distinct is needed to ensure, each ticket is only contained once
            allIssues.Where(i => i.Id.HasValue).Distinct().ForEach(i => issues[i.Id.Value] = i);
        }

        #endregion


        /// <summary>
        /// Class containing all and new projects
        /// </summary>
        public class ProjectLists
        {
            /// <summary>
            /// Gets or sets the list containing all known projects
            /// </summary>
            public IDictionary<int, ProjectInfo> AllProjects { get; set; }

            /// <summary>
            /// Gets or sets the list containing the new projects
            /// </summary>
            public List<ProjectInfo> NewProjects { get; set; }
        }

    }
}