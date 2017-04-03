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

    /// <summary>
    /// Class that stores information for the ribbon bar.
    /// </summary>
    public class SyncState
    {
        #region Fields

        /// <summary>
        /// Field for the <see cref="HoursInDay"/> property. It stores the amount of worked hours on the current day
        /// </summary>
        private double _hoursInDay;

        /// <summary>
        /// Field for the <see cref="HoursInMonth"/> property. It stores the amount of worked hours in the current month
        /// </summary>
        private double _hoursInMonth;

        /// <summary>
        /// Field for the <see cref="HoursInView"/> property. It stores the amount of worked hours that are visible in the current calendar view.
        /// </summary>
        private double _hoursInView;

        /// <summary>
        /// Field for the <see cref="HoursInWeek"/> property. It stores the amount of worked hours in the current week
        /// </summary>
        private double _hoursInWeek;

        /// <summary>
        /// Field for the <see cref="Status"/> property. Stores a string that indicates the synchronization state.
        /// </summary>
        private string _status;

        #endregion

        #region Public Events

        /// <summary>
        /// Event is raised when the connection state changed
        /// </summary>
        public event EventHandler ConnectionStateChanged;

        /// <summary>
        /// Event is raised when the status changed
        /// </summary>
        public event EventHandler StatusChanged;

        #endregion


        #region Public properties

        /// <summary>
        /// Gets or sets the status information
        /// </summary>
        public string Status
        {
            get
            {
                return this._status;
            }
            set
            {
                this._status = value;
                if (this.StatusChanged != null)
                {
                    this.StatusChanged(this, new EventArgs());
                }
            }
        }

        /// <summary>
        /// Gets or sets the hours in view
        /// </summary>
        public double HoursInView
        {
            get
            {
                return this._hoursInView;
            }
            set
            {
                this._hoursInView = value;
                if (this.StatusChanged != null)
                {
                    this.StatusChanged(this, new EventArgs());
                }
            }
        }

        /// <summary>
        /// Gets or sets the hours in view
        /// </summary>
        public double HoursInMonth
        {
            get
            {
                return this._hoursInMonth;
            }
            set
            {
                this._hoursInMonth = value;
                if (this.StatusChanged != null)
                {
                    this.StatusChanged(this, new EventArgs());
                }
            }
        }

        /// <summary>
        /// Gets or sets the hours in view
        /// </summary>
        public double HoursInWeek
        {
            get
            {
                return this._hoursInWeek;
            }
            set
            {
                this._hoursInWeek = value;
                if (this.StatusChanged != null)
                {
                    this.StatusChanged(this, new EventArgs());
                }
            }
        }

        /// <summary>
        /// Gets or sets the hours in view
        /// </summary>
        public double HoursInDay
        {
            get
            {
                return this._hoursInDay;
            }
            set
            {
                this._hoursInDay = value;
                if (this.StatusChanged != null)
                {
                    this.StatusChanged(this, new EventArgs());
                }
            }
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Raises the connection changed event
        /// </summary>
        public void RaiseConnectionChanged()
        {
            if (this.ConnectionStateChanged != null)
            {
                this.ConnectionStateChanged(this, new EventArgs());
            }
        }

        #endregion

    }
}