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

namespace Scorpio.Outlook.AddIn.UserInterface.ViewModel
{
    using System;
    using System.ComponentModel;
    using System.Linq;

    using Scorpio.Outlook.AddIn.Misc;

    /// <summary>
    /// The show time entries view model
    /// </summary>
    public class ShowTimeEntriesViewModel : ViewModelBase
    {
        #region Fields

        /// <summary>
        /// Internal member for the <see cref="BeginDate"/> Property
        /// </summary>
        private DateTime _beginDate;

        /// <summary>
        /// Internal member for the <see cref="BookedHoursInTimeSpan"/> Property
        /// </summary>
        private double _bookedHoursInTimeSpan;

        /// <summary>
        /// Internal member for the <see cref="EndDate"/> Property
        /// </summary>
        private DateTime _endDate;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ShowTimeEntriesViewModel"/> class.
        /// </summary>
        public ShowTimeEntriesViewModel()
        {
            this.EndDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            this.BeginDate = this.EndDate.AddMonths(-1);
            this.PropertyChanged += this.OnDatesChanges;
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets or sets the begin date.
        /// </summary>
        public DateTime BeginDate
        {
            get
            {
                return this._beginDate;
            }
            set
            {
                this.SetProperty(ref this._beginDate, value);
            }
        }

        /// <summary>
        /// Gets or sets the booked hoursin time span.
        /// </summary>
        public double BookedHoursInTimeSpan
        {
            get
            {
                return this._bookedHoursInTimeSpan;
            }
            set
            {
                this.SetProperty(ref this._bookedHoursInTimeSpan, value);
            }
        }

        /// <summary>
        /// Gets or sets the end date.
        /// </summary>
        public DateTime EndDate
        {
            get
            {
                return this._endDate;
            }
            set
            {
                this.SetProperty(ref this._endDate, value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Updates the <see cref="BookedHoursInTimeSpan"/> property when <see cref="BeginDate"/> or <see cref="EndDate"/> changes.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="propertyChangedEventArgs">The event args</param>
        private void OnDatesChanges(object sender, PropertyChangedEventArgs propertyChangedEventArgs)
        {
            if (new[] { nameof(this.BeginDate), nameof(this.EndDate) }.Contains(propertyChangedEventArgs.PropertyName))
            {
                this.BookedHoursInTimeSpan = Globals.ThisAddIn.UiUserInfoSynchronizer.GetHours(this.BeginDate, this.EndDate);
            }
        }

        #endregion
    }
}