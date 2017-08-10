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

namespace Scorpio.Outlook.AddIn.UserInterface.ViewModel
{
    using System;
    using System.ComponentModel;
    using System.Threading.Tasks;
    using System.Windows.Input;

    using Scorpio.Outlook.AddIn.Misc;
    using Scorpio.Outlook.AddIn.UserInterface.Helper;

    /// <summary>
    /// View model for the custom task pane.
    /// </summary>
    public class ScorpioTaskPaneViewModel : INotifyPropertyChanged
    {
        #region Fields

        /// <summary>
        /// Backing field for <see cref="ConnectString"/>
        /// </summary>
        private string _connectString;

        /// <summary>
        /// Backing field for <see cref="HoursCalendar"/>
        /// </summary>
        private double _hoursCalendar;

        /// <summary>
        /// Backing field for <see cref="HoursDay"/>
        /// </summary>
        private double _hoursDay;

        /// <summary>
        /// Backing field for <see cref="HoursMonth"/>
        /// </summary>
        private double _hoursMonth;

        /// <summary>
        /// Backing field for <see cref="HoursWeek"/>
        /// </summary>
        private double _hoursWeek;

        /// <summary>
        /// Backing field for <see cref="OpenCalendarCommand"/>
        /// </summary>
        private ICommand _openCalendarCommand;

        /// <summary>
        /// Backing field for <see cref="ResetTimeEntriesCommand"/>
        /// </summary>
        private ICommand _resetTimeEntriesCommand;

        /// <summary>
        /// Backing field for <see cref="SaveTimeEntriesCommand"/>
        /// </summary>
        private ICommand _saveTimeEntriesCommand;

        /// <summary>
        /// Backing field for <see cref="SynchronizeRedmineCommand"/>
        /// </summary>
        private ICommand _synchronizeRedmineCommand;
        
        /// <summary>
        /// The ui-context task scheduler.
        /// </summary>
        private TaskScheduler uiContext = TaskScheduler.FromCurrentSynchronizationContext();

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ScorpioTaskPaneViewModel"/> class.
        /// </summary>
        public ScorpioTaskPaneViewModel()
        {
                Globals.ThisAddIn.SyncState.StatusChanged += (sender, args) =>
                    {
                        this.HoursCalendar = Globals.ThisAddIn.SyncState.HoursInView;
                        this.HoursDay = Globals.ThisAddIn.SyncState.HoursInDay;
                        this.HoursWeek = Globals.ThisAddIn.SyncState.HoursInWeek;
                        this.HoursMonth = Globals.ThisAddIn.SyncState.HoursInMonth;
                        this.ConnectString = Globals.ThisAddIn.SyncState.Status;
                        var taskPaneTask = new Task(CommandManager.InvalidateRequerySuggested);
                        taskPaneTask.Start(this.uiContext);
                    };
        }

        #endregion

        #region Public Events

        /// <summary>
        /// Property changed event.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        #endregion

        #region Methods

        /// <summary>
        /// Method to raise a property changed event.
        /// </summary>
        /// <param name="name">The name of the property which changed</param>
        protected void OnPropertyChanged(string name)
        {
            var handler = this.PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }

        #endregion

        #region Public properties
        
        /// <summary>
        /// Gets or sets the string value which represents the current synchronization state.
        /// </summary>
        public string ConnectString
        {
            get
            {
                return this._connectString;
            }
            set
            {
                if (this._connectString == value)
                {
                    return;
                }
                this._connectString = value;
                this.OnPropertyChanged("ConnectString");
            }
        }

        /// <summary>
        /// Gets or sets the amount of worked hours that are in the currently opened calendar view range.
        /// </summary>
        public double HoursCalendar
        {
            get
            {
                return this._hoursCalendar;
            }
            set
            {
                if (Math.Abs(this._hoursCalendar - value) < Constants.Epsilon)
                {
                    return;
                }
                this._hoursCalendar = value;
                this.OnPropertyChanged("HoursCalendar");
            }
        }

        /// <summary>
        /// Gets or sets the amount of hours worked on the current day.
        /// </summary>
        public double HoursDay
        {
            get
            {
                return this._hoursDay;
            }
            set
            {
                if (Math.Abs(this._hoursDay - value) < Constants.Epsilon)
                {
                    return;
                }
                this._hoursDay = value;
                this.OnPropertyChanged("HoursDay");
            }
        }

        /// <summary>
        /// Gets or sets the amount of hours worked on the current day.
        /// </summary>
        public double HoursWeek
        {
            get
            {
                return this._hoursWeek;
            }
            set
            {
                if (Math.Abs(this._hoursWeek - value) < Constants.Epsilon)
                {
                    return;
                }
                this._hoursWeek = value;
                this.OnPropertyChanged("HoursWeek");
            }
        }

        /// <summary>
        /// Gets or sets the amount of hours worked in the current month.
        /// </summary>
        public double HoursMonth
        {
            get
            {
                return this._hoursMonth;
            }
            set
            {
                if (Math.Abs(this._hoursMonth - value) < Constants.Epsilon)
                {
                    return;
                }
                this._hoursMonth = value;
                this.OnPropertyChanged("HoursMonth");
            }
        }

        /// <summary>
        /// Gets a value indicating whether tickets and projects can currently be synchronized.
        /// </summary>
        public bool CanSynchronizeTickets
        {
            get
            {
                return !Globals.ThisAddIn.Synchronizer.IsConnecting;
            }
        }

        /// <summary>
        /// Gets a value indicating whether tickets can currently be synchronized.
        /// </summary>
        public bool CanSynchronizeTimeEntries
        {
            get
            {
                return Globals.ThisAddIn.Synchronizer.CanSyncTimeEntries;
            }
        }

        /// <summary>
        /// Gets the command which shows the redmine calendar.
        /// </summary>
        public ICommand OpenCalendarCommand
        {
            get
            {
                return this._openCalendarCommand
                       ?? (this._openCalendarCommand = new RelayCommand<object>(o => Globals.ThisAddIn.OpenCalendar(), o => true));
            }
        }
        
        /// <summary>
        /// Gets the command which synchronizes issue and project information with redmine.
        /// </summary>
        public ICommand SynchronizeRedmineCommand
        {
            get
            {
                return this._synchronizeRedmineCommand
                       ?? (this._synchronizeRedmineCommand =
                           new RelayCommand<object>(o => Globals.ThisAddIn.ReconnectToRedmine(), o => this.CanSynchronizeTickets));
            }
        }

        /// <summary>
        /// Gets the command which synchronizes time entries with redmine, saving changes made by the user.
        /// </summary>
        public ICommand SaveTimeEntriesCommand
        {
            get
            {
                return this._saveTimeEntriesCommand
                       ?? (this._saveTimeEntriesCommand =
                           new RelayCommand<object>(o => Globals.ThisAddIn.Synchronizer.SaveTimeEntriesAsync(), o => this.CanSynchronizeTimeEntries));
            }
        }

        /// <summary>
        /// Gets the command which resets the time-entries in the calendar to the state of the time entries in redmine, discarding unsaved changes.
        /// </summary>
        public ICommand ResetTimeEntriesCommand
        {
            get
            {
                return this._resetTimeEntriesCommand
                       ?? (this._resetTimeEntriesCommand =
                           new RelayCommand<object>(
                               o => Globals.ThisAddIn.Synchronizer.RevertTimeEntriesToRedmineState(),
                               o => this.CanSynchronizeTimeEntries));
            }
        }

        #endregion
    }
}