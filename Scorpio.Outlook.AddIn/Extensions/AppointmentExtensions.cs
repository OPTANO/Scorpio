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

namespace Scorpio.Outlook.AddIn.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using log4net;

    using Microsoft.Office.Interop.Outlook;

    using Redmine.Net.Api.Types;

    using Scorpio.Outlook.AddIn.Cache;
    using Scorpio.Outlook.AddIn.Helper;
    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Misc;
    using Scorpio.Outlook.AddIn.Properties;

    /// <summary>
    /// Class that contains extension methods to the <see cref="AppointmentItem"/> class.
    /// </summary>
    public static class AppointmentExtensions
    {
        #region Static Fields

        /// <summary>
        /// The logger.
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(AppointmentExtensions));

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Method to get the string representation of an appointment item
        /// </summary>
        /// <param name="item">the appointment</param>
        /// <returns>the string representation</returns>
        public static string GetStringRepresentation(this AppointmentItem item)
        {
            var stringRepresentation = string.Format(
                        "{0} - {1}, {2}: {3}",
                        item.Start,
                        item.End,
                        item.Location,
                        item.Subject);
            return stringRepresentation;
        }

        /// <summary>
        /// Appends a message to an appointment body
        /// </summary>
        /// <param name="item">The appointment item</param>
        /// <param name="message">The message</param>
        public static void AppendToBody(this AppointmentItem item, string message)
        {
            if (string.IsNullOrWhiteSpace(item.Body))
            {
                item.Body = message;
            }
            else
            {
                item.Body = item.Body + "\n" + message;
            }
        }

        /// <summary>
        /// Checks whether the appointment is different to the time entry
        /// </summary>
        /// <param name="item">The item</param>
        /// <param name="timeEntry">The time entry</param>
        /// <returns>True if the item is modified</returns>
        public static bool CheckItemIsModified(this AppointmentItem item, TimeEntryInfo timeEntry)
        {
            // create new calendar item
            var startTime = timeEntry.StartDateTime;
            var endTime = timeEntry.EndDateTime;

            if (!string.Equals(item.Subject, timeEntry.Name))
            {
                return true;
            }
            if (item.Start != startTime || item.End != endTime)
            {
                return true;
            }

            var issueId = item.GetAppointmentCustomId(Constants.FieldRedmineIssueId);
            if (issueId != timeEntry.IssueInfo.Id)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Creates the location string for an appointment
        /// </summary>
        /// <param name="item">The <see cref="AppointmentItem"/> for which to set the location.</param>
        /// <param name="issueId">The issue id</param>
        /// <param name="issue">The issue</param>
        public static void CreateAppointmentLocation(this AppointmentItem item, int issueId, IssueInfo issue)
        {
            item.Location = string.Format("#{0} - {1}", issueId, issue != null ? issue.Name : "???");
        }

        /// <summary>
        /// Gets a custom int field value from an appointment
        /// </summary>
        /// <param name="item">
        /// The appointment item
        /// </param>
        /// <param name="fieldName">
        /// The field name
        /// </param>
        /// <returns>
        /// The id or null if no value is set
        /// </returns>
        public static int? GetAppointmentCustomId(this AppointmentItem item, string fieldName)
        {
            var pid = item.UserProperties.Find(fieldName);
            if (pid == null || pid.Value == null || string.IsNullOrWhiteSpace(pid.Value.ToString()))
            {
                return null;
            }
            return Convert.ToInt32(pid.Value);
        }

        /// <summary>
        /// Gets the redmine modification timestamp of an item
        /// </summary>
        /// <param name="item">The appointment item</param>
        /// <returns>The modification date</returns>
        public static DateTime GetAppointmentModificationDate(this AppointmentItem item)
        {
            var prop = item.UserProperties.Find(Constants.FieldLastUpdate);
            if (prop == null || prop.Value == null)
            {
                return DateTime.MinValue;
            }

            return (DateTime)prop.Value;
        }

        /// <summary>
        /// Checks whether an appointment is copied in outlook
        /// </summary>
        /// <param name="item">The item to check</param>
        /// <returns>True if it is a copied item</returns>
        public static bool IsAppointmentCopied(this AppointmentItem item)
        {
            var prop = item.UserProperties.Find(Constants.FieldEntryIdCopy);
            if (prop == null)
            {
                return false;
            }

            return !string.Equals(prop.Value.ToString(), item.EntryID);
        }

        /// <summary>
        /// Checks whether a category is set for an item
        /// </summary>
        /// <param name="item">The appointment item</param>
        /// <param name="appointmentCategory">The category for which to check if it is set</param>
        /// <returns>True if the category is set</returns>
        public static bool IsCategorySet(this AppointmentItem item, string appointmentCategory)
        {
            // get the previous categories and remove our categories
            if (!string.IsNullOrEmpty(item.Categories))
            {
                foreach (var itm in item.Categories.Split(';'))
                {
                    if (itm == appointmentCategory)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Checks whether the deleted category is set for an item
        /// </summary>
        /// <param name="item">The appointment item</param>
        /// <returns>True if the category is set</returns>
        public static bool IsDeletedSet(this AppointmentItem item)
        {
            return item.IsCategorySet(AppointmentState.Deleted.Name);
        }

        /// <summary>
        /// Checks whether the modified category is set for an item
        /// </summary>
        /// <param name="item">The appointment item</param>
        /// <returns>True if the category is set</returns>
        public static bool IsModifiedSet(this AppointmentItem item)
        {
            return item.IsCategorySet(AppointmentState.Modified.Name);
        }

        /// <summary>
        /// Checks whether the sync error category is set for an item
        /// </summary>
        /// <param name="item">The appointment item</param>
        /// <returns>True if the category is set</returns>
        public static bool IsSyncErrorSet(this AppointmentItem item)
        {
            return item.IsCategorySet(AppointmentState.SyncError.Name);
        }

        /// <summary>
        /// Checks whether the synchronized overtime category is set for an item
        /// </summary>
        /// <param name="item">The appointment item</param>
        /// <returns>True if the category is set</returns>
        public static bool IsSynchronizedOvertimeSet(this AppointmentItem item)
        {
            return item.IsCategorySet(AppointmentState.SynchronizedOvertime.Name);
        }

        /// <summary>
        /// Checks whether the synchronized category is set for an item
        /// </summary>
        /// <param name="item">The appointment item</param>
        /// <returns>True if the category is set</returns>
        public static bool IsSynchronizedSet(this AppointmentItem item)
        {
            return item.IsCategorySet(AppointmentState.Synchronized.Name);
        }

        /// <summary>
        /// Mark the item as not copied
        /// </summary>
        /// <param name="item">The appointment item</param>
        public static void MarkAsNotCopied(this AppointmentItem item)
        {
            var field = OutlookHelper.CreateOrGetProperty(item, Constants.FieldEntryIdCopy, OlUserPropertyType.olText);
            field.Value = item.EntryID;
            item.Save();
        }

        /// <summary>
        /// Method that checks the previous state of an appointment and tries to reset that state.
        /// </summary>
        /// <param name="item">The appointment for which to set the previous state.</param>
        public static void ResetPreviousState(this AppointmentItem item)
        {
            var previousState = item.GetAppointmentCustomId(Constants.FieldAppointmentPreviousState);

            if (!previousState.HasValue)
            {
                return;
            }

            var state = AppointmentState.AllStates.FirstOrDefault(s => s.Value == previousState.Value);
            if (state == null)
            {
                return;
            }
            item.SetAppointmentState(state);
            item.Save();
        }

        /// <summary>
        /// Sets a custom field id for an appointment
        /// </summary>
        /// <param name="item">The item to set</param>
        /// <param name="fieldName">The field name</param>
        /// <param name="value">The value</param>
        public static void SetAppointmentCustomId(this AppointmentItem item, string fieldName, int? value)
        {
            var property = OutlookHelper.CreateOrGetProperty(item, fieldName, OlUserPropertyType.olInteger);
            property.Value = value;
        }

        /// <summary>
        /// Sets the last modification date of an appointment item
        /// </summary>
        /// <param name="item">The item</param>
        /// <param name="modDate">The modification date</param>
        public static void SetAppointmentModificationDate(this AppointmentItem item, DateTime modDate)
        {
            OutlookHelper.CreateOrGetProperty(item, Constants.FieldLastUpdate, OlUserPropertyType.olDateTime).Value = modDate;
        }

        /// <summary>
        /// Sets the appointment category to give the appointment the correct color
        /// </summary>
        /// <param name="item">The item to set</param>
        /// <param name="state">The appointment state to set to the appointment</param>
        public static void SetAppointmentState(this AppointmentItem item, AppointmentState state)
        {
            var newCategories = new List<string>();

            // get the previous categories and remove our categories
            if (!string.IsNullOrEmpty(item.Categories))
            {
                foreach (var itm in item.Categories.Split(';'))
                {
                    // Preserve other categories.
                    if (!AppointmentState.IsValidAppointmentStateName(itm))
                    {
                        newCategories.Add(itm);
                    }
                }
            }

            // set our new categories
            newCategories.Insert(0, state.Name);

            if (state == AppointmentState.Modified)
            {
                item.SetAppointmentModificationDate(DateTime.Now);
            }

            item.Categories = string.Join(";", newCategories);

            // Update the previous sate of the appointment, if the current state is not syncerror. 
            var previousState = item.GetAppointmentCustomId(Constants.FieldAppointmentState);
            if (previousState.HasValue && previousState.Value != AppointmentState.SyncError.Value)
            {
                item.SetAppointmentCustomId(Constants.FieldAppointmentPreviousState, previousState);
            }

            item.SetAppointmentCustomId(Constants.FieldAppointmentState, state.Value);
        }

        /// <summary>
        /// Updates the custom fields of an appointment
        /// </summary>
        /// <param name="appointment">
        /// The appointment
        /// </param>
        /// <param name="timeEntryId">
        /// The redmine time entry id
        /// </param>
        /// <param name="projectId">
        /// The redmine project id
        /// </param>
        /// <param name="issueId">
        /// The redmine issue id
        /// </param>
        /// <param name="activityId">
        /// The activity Id.
        /// </param>
        /// <param name="lastRedmineUpdate">
        /// The date/time when the item was updated in redmine the last time
        /// </param>
        public static void UpdateAppointmentFields(
            this AppointmentItem appointment, 
            int? timeEntryId, 
            int projectId, 
            int issueId, 
            int activityId, 
            DateTime lastRedmineUpdate)
        {
            // create user properties
            if (timeEntryId.HasValue)
            {
                appointment.SetAppointmentCustomId(Constants.FieldRedmineTimeEntryId, timeEntryId);
            }
            appointment.SetAppointmentCustomId(Constants.FieldRedmineProjectId, projectId);
            appointment.SetAppointmentCustomId(Constants.FieldRedmineIssueId, issueId);
            appointment.SetAppointmentCustomId(Constants.FieldRedmineActivityId, activityId);

            // create last update field
            var propLastUpdate = OutlookHelper.CreateOrGetProperty(appointment, Constants.FieldLastUpdate, OlUserPropertyType.olDateTime);
            propLastUpdate.Value = lastRedmineUpdate;
        }

        /// <summary>
        /// Updates an appointment to the values of a redmine time entry
        /// </summary>
        /// <param name="item">The appointment item</param>
        /// <param name="timeEntry">The redmine time entry</param>
        /// <param name="issue">The issue that belongs to the timeentry. Can be <code>null</code> if the issue is not known.</param>
        public static void UpdateAppointmentFromTimeEntry(this AppointmentItem item, TimeEntryInfo timeEntry, IssueInfo issue)
        {
            // create new calendar item
            item.CreateAppointmentLocation(timeEntry.IssueInfo.Id, issue);
            item.Subject = timeEntry.Name;
            item.Start = timeEntry.StartDateTime;
            var endTime = timeEntry.EndDateTime;

            if (timeEntry.IssueInfo.Id != Settings.Default.RedmineUseOvertimeIssue)
            {
                item.End = timeEntry.EndDateTime;
            }
            else
            {
                // In case of the overtime issue, we cannot take the actual hours property for determining the end time, because hours is 0.
                item.End = endTime;
            }

            if (timeEntry.IssueInfo.Id != Settings.Default.RedmineUseOvertimeIssue && Math.Abs((endTime - item.End).TotalMinutes) > 5)
            {
                item.AppendToBody("Die Von/Bis-Zeit dieses Elements wurde geändert, da sie nicht mit dem 'Stunden'-Feld aus Redmine übereinstimmte.");
                Log.WarnFormat("The end time of the time entry was changed, because it was not consistent with the provided duration. ");
            }

            // create user properties
            item.UpdateAppointmentFields(timeEntry.Id, timeEntry.ProjectInfo.Id, timeEntry.IssueInfo.Id, timeEntry.ActivityInfo.Id, timeEntry.UpdateTime);

            if (timeEntry.IssueInfo.Id != Settings.Default.RedmineUseOvertimeIssue)
            {
                item.SetAppointmentState(AppointmentState.Synchronized);
            }
            else
            {
                item.SetAppointmentState(AppointmentState.SynchronizedOvertime);
            }
        }

        #endregion
    }
}