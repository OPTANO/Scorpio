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

namespace Scorpio.Outlook.AddIn.Helper
{
    using System;
    using System.Linq;

    /// <summary>
    /// Helper class for interaction with <see cref="DateTime"/> objects.
    /// </summary>
    public static class DateTimeHelper
    {
        #region Constants
        
        /// <summary>
        /// The sunday to previous monday offset.
        /// </summary>
        private const int SundayToPreviousMondayOffset = -6;

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Method that checks if a given date is a workday.
        /// </summary>
        /// <param name="date">The date for which to check whether it is a workday.</param>
        /// <returns><code>false</code> if the date is a saturday or sunday, <code>true</code> otherwise.</returns>
        public static bool IsWorkDay(DateTime date)
        {
            return date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday;
        }
        
        /// <summary>
        /// Gets the start date of the week
        /// </summary>
        /// <param name="date">the date to check</param>
        /// <returns>the date of the monday of the week</returns>
        public static DateTime StartOfWeek(DateTime date)
        {
            var weekDate = date.Date;
            var dayOffset = weekDate.DayOfWeek == DayOfWeek.Sunday ? SundayToPreviousMondayOffset : (int)DayOfWeek.Monday - (int)weekDate.DayOfWeek;
            var startOfWeek = weekDate.AddDays(dayOffset);
            return startOfWeek;
        }

        /// <summary>
        /// Gets the start day of the next/following week
        /// </summary>
        /// <param name="date">the date to check</param>
        /// <returns>the date of the sunday of teh week</returns>
        public static DateTime StartOfNextWeek(DateTime date)
        {
            var startOfWeek = StartOfWeek(date);
            var startOfNextWeek = startOfWeek.AddDays(7);

            return startOfNextWeek;
        }
        
        /// <summary>
        /// Gets the start date of the month
        /// </summary>
        /// <param name="date">the date to check</param>
        /// <returns>the start date of the month</returns>
        public static DateTime StartOfMonth(DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
        }

        /// <summary>
        /// Gets the start date of the next/following month
        /// </summary>
        /// <param name="date">the date to check</param>
        /// <returns>the end date</returns>
        public static DateTime StartOfNextMonth(DateTime date)
        {
            var startOfNextMonth = StartOfMonth(date).AddMonths(1);
            return startOfNextMonth;
        }

        /// <summary>
        /// Gets the start of the day
        /// </summary>
        /// <param name="dateTime">the date and time of interest</param>
        /// <returns>the date and time of the start of the day</returns>
        public static DateTime StartOfDay(DateTime dateTime)
        {
            return dateTime.Date;
        }

        /// <summary>
        /// Gets the start of the next/following day
        /// </summary>
        /// <param name="dateTime">the date and time of interest</param>
        /// <returns>the date and time of the end of the day</returns>
        public static DateTime StartOfNextDay(DateTime dateTime)
        {
            return dateTime.Date.AddDays(1).Date;
        }

        /// <summary>
        /// The minimum date time
        /// </summary>
        /// <param name="dates">the dates to check</param>
        /// <returns>the minimum value, date time max, if no values are given</returns>
        public static DateTime MinimumDateTime(params DateTime[] dates)
        {
            return dates.DefaultIfEmpty(DateTime.MaxValue).Min();
        }

        /// <summary>
        /// The maximum date time
        /// </summary>
        /// <param name="dates">the dates to check</param>
        /// <returns>the maximum value, date time min, if no values are given</returns>
        public static DateTime MaximumDateTime(params DateTime[] dates)
        {
            return dates.DefaultIfEmpty(DateTime.MinValue).Max();
        }

        #endregion
    }
}