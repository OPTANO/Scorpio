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

    using Scorpio.Outlook.AddIn.Misc;

    using Exception = System.Exception;

    /// <summary>
    /// Extension Methods for the <see cref="MAPIFolder"/> class.
    /// </summary>
    public static class MapiFolderExtensions
    {
        #region Static Fields

        /// <summary>
        /// The logger.
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(MapiFolderExtensions));

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Gets the filter string to use for querying the appointments intersecting with the given time range. If start/end are included
        /// depends on the parameter set.
        /// </summary>
        /// <param name="startTime">Start of the date range</param>
        /// <param name="endTime">End of the date range</param>
        /// <param name="includeStart">if the start is included or not</param>
        /// <param name="includeEnd">if the end is included or not</param>
        /// <returns>Query string for getting appointment items in the date range</returns>
        internal static string GetFilterString(DateTime startTime, DateTime endTime, bool includeStart = true, bool includeEnd = true)
        {
            var filter = string.Format(
                "[Start] {2} '{0:g}' AND [End] {3} '{1:g}'",
                endTime,
                startTime,
                includeEnd ? "<=" : "<",
                includeStart ? ">=" : ">");
            
            return filter;
        }

        /// <summary>
        /// Get recurring appointments in date range.
        /// </summary>
        /// <param name="folder">The <see cref="MAPIFolder"/> from which to get the appointments.</param>
        /// <param name="startTime">Start of the date range</param>
        /// <param name="endTime">End of the date range</param>
        /// <param name="includeStart">if the start is included or not</param>
        /// <param name="includeEnd">if the end is included or not</param>
        /// <returns>Outlook appointment items in the date range, sorted by start date-time</returns>
        public static List<AppointmentItem> GetAppointmentsInRange(this MAPIFolder folder, DateTime startTime, DateTime endTime, bool includeStart = true, bool includeEnd = true)
        {
            var filter = GetFilterString(startTime, endTime, includeStart, includeEnd);

            try
            {
                var calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                var restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                {
                    return restrictItems.OfType<AppointmentItem>().ToList();
                }
                else
                {
                    return new List<AppointmentItem>();
                }
            }
            catch (Exception ex)
            {
                Log.Error(string.Format("Could not filter for appointments in time range with filter {0}.", filter), ex);
                return null;
            }
        }

        /// <summary>
        /// Gets all <see cref="AppointmentItem"/> from a <see cref="MAPIFolder"/> that are modified. An appointment is modified if it has 
        /// the userproperty <see cref="Constants.FieldAppointmentState"/> set to the value of <see cref="AppointmentState.Deleted"/>
        /// or <see cref="AppointmentState.Modified"/> or <see cref="AppointmentState.SyncError"/>. 
        /// </summary>
        /// <param name="folder">The folder from which to get the modified elements.</param>
        /// <returns>The modified appointment items in that folder.</returns>
        public static List<AppointmentItem> GetAppointmentsWithModification(this MAPIFolder folder)
        {
            var filter = "[" + Constants.FieldAppointmentState + "] = " + AppointmentState.Deleted.Value + " OR [" + Constants.FieldAppointmentState
                         + "] = " + AppointmentState.Modified.Value + " OR [" + Constants.FieldAppointmentState + "] = "
                         + AppointmentState.SyncError.Value;
                
            try
            {
                var calItems = folder.Items;
                calItems.IncludeRecurrences = true;

                // calItems.Sort("[Start]", Type.Missing);
                var restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                {
                    return restrictItems.OfType<AppointmentItem>().ToList();
                }

                return new List<AppointmentItem>();
            }
            catch (Exception ex)
            {
                Log.Error("Could not filter for appointments with modifications: ", ex);
                return null;
            }
        }

        #endregion
    }
}