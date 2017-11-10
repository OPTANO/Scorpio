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
    /// Interface implemented by an external data source, i.e. redmine that is able to write time entries and return the current ones.
    /// </summary>
    public interface IExternalSource
    {
        /// <summary>
        /// Gets or sets the limit for a request made
        /// </summary>
        int Limit { get; set; }

        /// <summary>
        /// Gets a hash set containing all project ids whose status is watched
        /// </summary>
        HashSet<int> ProjectsWithWatchedIssueStatus { get; }

            /// <summary>
        /// Method to create an object from teh given time entry
        /// </summary>
        /// <param name="entry">the time entry</param>
        /// <exception cref="CrudException">this exception is thrown if the object could not be created in the external source and must be handled by the calling method</exception>
        /// <exception cref="ConnectionException">raised if it is not possible to establish a connection to the external source</exception>
        /// <returns>the object created</returns>
        TimeEntryInfo CreateObject(TimeEntryInfo entry);

        /// <summary>
        /// Method to update the given object with the data contained in the object. The object has the id set, all other values should be updated
        /// </summary>
        /// <param name="entry">the entry to update</param>
        /// <returns>the time entry info of the updated object</returns>
        TimeEntryInfo UpdateObject(TimeEntryInfo entry);

        /// <summary>
        /// Method to delete an object by its name
        /// </summary>
        /// <param name="timeEntryId">the id of the time entry, if id is null, nothing should be deleted</param>
        /// <param name="nameValueCollection">the name value collection</param>
        /// <exception cref="ConnectionException">thrown if the conncection couldnot be established</exception>
        void DeleteTimeEntry(int? timeEntryId, DataSourceParameter nameValueCollection);

        /// <summary>
        /// Method to get the current user logged in to redmine
        /// </summary>
        /// <returns>the user</returns>
        UserInfo GetCurrentUser();

        /// <summary>
        /// Method to get the total object list of all projects
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        IList<ProjectInfo> GetTotalProjectList(DataSourceParameter parameters, Action<int, int> statusCallback = null);

        /// <summary>
        /// Method to get the total object list of all issues
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        IList<IssueInfo> GetTotalIssueInfoList(DataSourceParameter parameters, Action<int, int> statusCallback = null);

        /// <summary>
        /// Method to get the list of all issues matching the given parameter
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <returns>the list containing all objects</returns>
        IList<IssueInfo> GetIssueInfoList(DataSourceParameter parameters);

        /// <summary>
        /// Method to get the total object list of all activities
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        IList<ActivityInfo> GetTotalActivityInfoList(DataSourceParameter parameters, Action<int, int> statusCallback = null);

        /// <summary>
        /// Method to get the total object list of all time entries
        /// </summary>
        /// <param name="parameters">the parameters</param>
        /// <param name="statusCallback">a statusCallback to be run after the list is obtained</param>
        /// <returns>the list containing all objects</returns>
        IList<TimeEntryInfo> GetTotalTimeEntryInfoList(DataSourceParameter parameters, Action<int, int> statusCallback = null);
    }
}