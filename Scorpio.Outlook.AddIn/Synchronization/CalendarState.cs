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

namespace Scorpio.Outlook.AddIn.Synchronization
{
    using System;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Threading;

    using log4net;

    using Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Class that keeps the necessary information about the calendar view in outlook.
    /// </summary>
    public class CalendarState
    {
        #region Static Fields

        /// <summary>
        /// The logger.
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(CalendarState));

        #endregion

        #region Fields

        /// <summary>
        /// The <see cref="CalendarView"/> which is visible.
        /// </summary>
        private CalendarView _calendarView;

        /// <summary>
        /// The timer which checks if the view changed.
        /// </summary>
        private Timer _checkViewTimer;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CalendarState"/> class.
        /// </summary>
        public CalendarState()
        {
            this._checkViewTimer = new System.Threading.Timer(this.CheckCurrentViewCallback, null, 100, 1000);
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Gets or sets the calendar view which is currently visible.
        /// </summary>
        public CalendarView CalendarView
        {
            get
            {
                return this._calendarView;
            }
            set
            {
                this._calendarView = value;

                // Raise the connection changed because it will invalidate the controls in the ribbon bar.
                // TODO: add another event, or rename.
                Globals.ThisAddIn.SyncState.RaiseConnectionChanged();
            }
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Gets the current month in view
        /// </summary>
        /// <returns>The current month or <see cref="DateTime.MinValue"/> if nothing is selected.</returns>
        public DateTime GetCurrentMonth()
        {
            var dates = this.GetDisplayDates();
            if (dates == null || !dates.Any())
            {
                return DateTime.MinValue;
            }
            var minDate = dates.Min();
            return new DateTime(minDate.Year, minDate.Month, 1);
        }

        /// <summary>
        /// Method to get the array of displayed dates from the currently open calendar view. If now calendar view is open, null is returned.
        /// </summary>
        /// <returns>The dates displayed in the current calendar view. <code>null</code> if there is no calendar view opened.</returns>
        public DateTime[] GetDisplayDates()
        {
            if (this.CalendarView == null)
            {
                return null;
            }
            return this.CalendarView.DisplayedDates as DateTime[];
        }

        #endregion

        #region Methods

        /// <summary>
        /// Callback to check the current view
        /// </summary>
        /// <param name="state">The timer state</param>
        private void CheckCurrentViewCallback(object state)
        {
            try
            {
                var currentExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (currentExplorer == null)
                {
                    return;
                }

                var calendarView = currentExplorer.CurrentView as CalendarView;
                var isOpen = calendarView != null;
                var wasOpen = this.CalendarView != null;
                if (isOpen != wasOpen)
                {
                    this.CalendarView = calendarView;
                }
            }
            catch (COMException ex)
            {
                Log.Error("COM Exception in callback - problably nothing serious, might occur when other operations need too long", ex);
            }
        }

        #endregion
    }
}