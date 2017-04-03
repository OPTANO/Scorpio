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

namespace Scorpio.Outlook.AddIn.Synchronization
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Timers;
    using System.Windows.Forms;

    using Microsoft.Office.Interop.Outlook;

    using Scorpio.Outlook.AddIn.Extensions;
    using Scorpio.Outlook.AddIn.Helper;
    using Scorpio.Outlook.AddIn.Misc;
    using Scorpio.Outlook.AddIn.Properties;

    using Timer = System.Timers.Timer;

    /// <summary>
    /// This class keeps hold that the information displayed to the user in the ui (additional infos) are kept up to date.
    /// This logic former was contained in the redmine synchronizer, but is now done in a separate class.
    /// </summary>
    public class UiUserInfoSynchronizer
    {
        #region Static Fields
        
        /// <summary>
        /// The ticket id of the overtime issue
        /// </summary>
        private static readonly int OverTimeIssueId = Settings.Default.RedmineUseOvertimeIssue;

        #endregion

        #region Fields

        /// <summary>
        /// Function to get the appointments in the given time range.
        /// </summary>
        private readonly Func<DateTime, DateTime, List<AppointmentItem>> _appointmentsInRangeFunction;

        /// <summary>
        /// The update time, triggering the ui updates
        /// </summary>
        private Timer _updateTimer;

        /// <summary>
        /// The last maximum date displayed in the ui
        /// </summary>
        private DateTime? _maxDateDisplayed;

        /// <summary>
        /// The last minimum date displayed in the ui
        /// </summary>
        private DateTime? _minDateDisplayed;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="UiUserInfoSynchronizer"/> class.
        /// </summary>
        /// <param name="getAppointmentsInRangeFunction">method to get the appointments in the given time range</param>
        public UiUserInfoSynchronizer(Func<DateTime, DateTime, List<AppointmentItem>> getAppointmentsInRangeFunction)
        {
            // set the get appointment function
            this._appointmentsInRangeFunction = getAppointmentsInRangeFunction;

            // start the timer
            this.RestartTimer();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Methot to start/restart the timer. If a timer already exists, the old timer is stopped and cleared, the new timer is started.
        /// </summary>
        /// <param name="newRefreshTime">the new refresh time to be set</param>
        public void RestartTimer(double? newRefreshTime = null)
        {
            // stop existing time
            if (this._updateTimer != null)
            {
                this._updateTimer.Stop();
                this._updateTimer.Close();
                this._updateTimer.Elapsed -= this.UpdateTimerCallback;
                this._updateTimer.SynchronizingObject = null;
                this._updateTimer = null;
            }

            // get update time span
            if (newRefreshTime != null)
            {
                // set the value to at least one second
                Settings.Default.RefreshTime = Math.Max(1, newRefreshTime.Value);
            }
            var uiRefreshSettingInSeconds = Settings.Default.RefreshTime;
            var refreshInMilliseconds = TimeSpan.FromSeconds(uiRefreshSettingInSeconds).TotalMilliseconds;
            
            // initialize timer
            this._updateTimer = new Timer(refreshInMilliseconds);
            this._updateTimer.Elapsed += this.UpdateTimerCallback;
            this._updateTimer.AutoReset = true;
            this._updateTimer.SynchronizingObject = new Control();
            this._updateTimer.Enabled = true;
        }

        /// <summary>
        /// Gets the total amount of hours from a list of appointment items. Appointment items that are for the overtime issue are not counted. 
        /// </summary>
        /// <param name="appointmentItems">The appointment items.</param>
        /// <param name="predicate">an additional filter predicate to filter the list of appointments, default is null, i.e. no further filter operations needed</param>
        /// <returns>The total amount of hours.</returns>
        private static double GetWorkTimeInHours(List<AppointmentItem> appointmentItems, Predicate<AppointmentItem> predicate = null)
        {
            // build the filter to use for filtering the appointments
            var predicateToUse = new Predicate<AppointmentItem>(
                                     ai =>
                                         {
                                             var issueId = ai.GetAppointmentCustomId(Constants.FieldRedmineIssueId);
                                             var value = issueId.HasValue && issueId != OverTimeIssueId;
                                             if (value)
                                             {
                                                 value = predicate == null || predicate(ai);
                                             }
                                             return value;
                                         });

            // calculate and return the working time in hours for all the appointments matching the filter criteria
            var workingTimeInMinutes = appointmentItems.Where(app => predicateToUse(app)).Select(app => app.Duration).DefaultIfEmpty(0).Sum();
            var workingTimeInHours = workingTimeInMinutes / (double)Constants.MinutesInHour;
            return workingTimeInHours;
        }

        /// <summary>
        /// Method to get a predicate for filtering appointment items due to times
        /// </summary>
        /// <param name="start">the start for the filtering time range</param>
        /// <param name="end">the end of the filtering time range</param>
        /// <param name="includeStart">if the start should be included</param>
        /// <param name="includeEnd">if the end should be included</param>
        /// <returns>the predicate for filtering the time range</returns>
        private Predicate<AppointmentItem> GetTimeFilterPredicate(DateTime start, DateTime end, bool includeStart = true, bool includeEnd = true)
        {
            var predicate = new Predicate<AppointmentItem>(
                                ai =>
                                    {
                                        var startOfApp = ai.Start;
                                        var endOfApp = ai.End;

                                        var startFits = includeStart ? end >= startOfApp : end > startOfApp;
                                        var endFits = includeEnd ? start <= endOfApp : start < endOfApp;

                                        return startFits && endFits;
                                    });

            return predicate;
        }

        /// <summary>
        /// Updates the hours status for the current view
        /// </summary>
        /// <param name="from">Start date</param>
        /// <param name="to">End date</param>
        private void UpdateHoursInView(DateTime from, DateTime to)
        {
            try
            {
                // This is ugly, because it will frequently get updated by a timer, regardless of whether something has actually changed. 
                // This will be changed in a later version, by employing change events that are fired whenever a timeentry is created/updated/revmoed/etc.
                var now = DateTime.Now;

                // get start points
                var startView = from;
                var startDay = DateTimeHelper.StartOfDay(now);
                var startOfWeek = DateTimeHelper.StartOfWeek(now);
                var startOfMonth = DateTimeHelper.StartOfMonth(now);

                // get end points
                var endView = DateTimeHelper.StartOfNextDay(to);
                var startOfNextDay = DateTimeHelper.StartOfNextDay(now);
                var startOfNextWeek = DateTimeHelper.StartOfNextWeek(now);
                var startOfNextMonth = DateTimeHelper.StartOfNextMonth(now);

                // get time ranges to query
                var maximumEnd = DateTimeHelper.MaximumDateTime(startOfNextMonth, startOfNextWeek);
                var minimumStart = DateTimeHelper.MinimumDateTime(startView, startOfMonth, startOfWeek);

                // get the appointments in the maximum time range
                var appointmentsInMaximumTimeRange = this._appointmentsInRangeFunction(minimumStart, maximumEnd);

                // get the appointments in the current time range, reuse the maximum ones, if the time ranges overlap
                // this is done due to saving queries of the appointments
                List<AppointmentItem> currentViewAppointments;
                if (startView >= minimumStart && endView <= maximumEnd)
                {
                    currentViewAppointments = appointmentsInMaximumTimeRange;
                }
                else
                {
                    currentViewAppointments = this._appointmentsInRangeFunction(startView, endView);
                }

                // update hours in view
                var hoursInView = GetWorkTimeInHours(currentViewAppointments, this.GetTimeFilterPredicate(startView, endView, includeEnd: false));
                Globals.ThisAddIn.SyncState.HoursInView = hoursInView;

                // update hours in current day
                var hoursInDay = GetWorkTimeInHours(
                    appointmentsInMaximumTimeRange,
                    this.GetTimeFilterPredicate(startDay, startOfNextDay, includeEnd: false));
                Globals.ThisAddIn.SyncState.HoursInDay = hoursInDay;

                // update hours in current week
                var hoursInWeek = GetWorkTimeInHours(
                    appointmentsInMaximumTimeRange,
                    this.GetTimeFilterPredicate(startOfWeek, startOfNextWeek, includeEnd: false));
                Globals.ThisAddIn.SyncState.HoursInWeek = hoursInWeek;

                // update hours in current month
                var hoursInMonth = GetWorkTimeInHours(
                    appointmentsInMaximumTimeRange,
                    this.GetTimeFilterPredicate(startOfMonth, startOfNextMonth, includeEnd: false));
                Globals.ThisAddIn.SyncState.HoursInMonth = hoursInMonth;
            }
            catch (System.Exception e)
            {

            }
        }

        /// <summary>
        /// Method which is called periodically by the <see cref="_updateTimer"/>. It keeps the display of 
        /// logged hours in the ribbon bar synchronized with the calendar view.
        /// The method checks if the dates displayed have changed, a recalculation of the worked hours in only performed,
        /// if there were changes in the dates displayed.
        /// The recalculation is also triggered after an appointment is added, changed or deleted. The periodically check if
        /// only done to ensure, that the values are updated after another calendar area is chosen, because there is no
        /// callback for that in outlook
        /// </summary>
        /// <param name="sender">The timer</param>
        /// <param name="args">The elapsed event args</param>
        private void UpdateTimerCallback(object sender, ElapsedEventArgs args)
        {
            /*
             * TODO: We should try to get rid of the update timer. However, Outlook does not seem 
             * to provide a callback for change of the displayed dates in the calendar view:
             * http://stackoverflow.com/questions/32693475/outlook-2013-vsto-get-calendar-selected-range-callback
             */

            // get dates and check if there are any selected, else nothing to do here
            var dates = Globals.ThisAddIn.CalendarState.GetDisplayDates();
            if (dates == null || dates.Length == 0)
            {
                return;
            }

            // get min and max date and compare them to the last min and max date stored
            var currentMinDate = dates.Min();
            var currentMaxDate = dates.Max();

            if (object.Equals(currentMinDate, this._minDateDisplayed) && object.Equals(currentMaxDate, this._maxDateDisplayed))
            {
                // nothing changed, nothing to do
            }
            else
            {
                // at least one of the two dates changed, update their values and recalculate
                this._minDateDisplayed = currentMinDate;
                this._maxDateDisplayed = currentMaxDate;

                // call the method to update the values in the ui
                this.UpdateHoursInView(this._minDateDisplayed.Value, this._maxDateDisplayed.Value);
            }
        }
        
        #endregion

        /// <summary>
        /// Method called after an appointment has been changed 
        /// </summary>
        /// <param name="sender">the sender</param>
        /// <param name="args">teh event arguments</param>
        public void HandleAppointmentChange(object sender, EventArgs args)
        {
            // get dates and check if there are any selected, else nothing to do here
            var dates = Globals.ThisAddIn.CalendarState.GetDisplayDates();
            if (dates == null || dates.Length == 0)
            {
                return;
            }

            // get min and max date and compare them to the last min and max date stored
            this.UpdateHoursInView(dates.Min(), dates.Max());
        }
    }
}