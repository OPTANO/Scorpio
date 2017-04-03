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

namespace Scorpio.Outlook.AddIn.UserInterface.Controls
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using System.Windows;

    using Scorpio.Outlook.AddIn.Cache;
    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Misc;

    /// <summary>
    /// Interaction logic for RecurringTimeEntryDialog.xaml
    /// </summary>
    public partial class RecurringTimeEntryDialog : INotifyPropertyChanged
    {
        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="RecurringTimeEntryDialog"/> class.
        /// </summary>
        public RecurringTimeEntryDialog()
        {
            this.InitializeComponent();
            this.DataContext = this;
        }

        #endregion

        #region Public Events

        /// <summary>
        /// The property changed event.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets or sets a list of available issues for selection. TODO DS: Maybe sort Abwesenheit Issues to the front
        /// </summary>
        public List<IssueInfo> AvailableIssues { get; set; }

        /// <summary>
        /// Gets or sets the description comment of the recurring time entry.
        /// </summary>
        public string Description { get; set; } = "Automatische Serienbuchung";

        /// <summary>
        /// Gets or sets the date at which the last booking for the recurring time entries should occur.
        /// </summary>
        public DateTime EndDate { get; set; } = DateTime.Today.Date;

        /// <summary>
        /// Gets or sets the time of day at which each time entry of the series ends.
        /// </summary>
        public DateTime EndTime { get; set; } = DateTime.Today.Date.AddHours(17);

        /// <summary>
        /// Gets or sets a value indicating whether there should be recurring time entries on weekends for this series.
        /// </summary>
        public bool IsBookingOnWeekends { get; set; } = false;

        /// <summary>
        /// Gets or sets the issue for which recurring time entries are created.
        /// </summary>
        public IssueInfo SelectedIssue { get; set; }

        /// <summary>
        /// Gets or sets the first day for which to create a recurring time entry.
        /// </summary>
        public DateTime StartDate { get; set; } = DateTime.Today.Date;

        /// <summary>
        /// Gets or sets the time of the day at which the recurring time entries should begin.
        /// </summary>
        public DateTime StartTime { get; set; } = DateTime.Today.Date.AddHours(9);

        /// <summary>
        /// Gets or sets a list of validation messages that indicate errors in the configuration for the recurring time entries.
        /// </summary>
        public List<string> ValidationMessages { get; set; } = new List<string>();

        /// <summary>
        /// Gets the validation messages as a single string.
        /// </summary>
        public string ValidationMessagesString
        {
            get
            {
                if (this.ValidationMessages.Any())
                {
                    return this.ValidationMessages.Aggregate((current, next) => current + Environment.NewLine + next);
                }

                return "";
            }
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Cancels the dialog, indicating that no recurring time entries should be created.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="routedEventArgs">The event arguments</param>
        public void CancelClicked(object sender, RoutedEventArgs routedEventArgs)
        {
            this.DialogResult = false;
            this.Close();
        }

        /// <summary>
        /// Cancels the dialog, indicating that the recurring time entries should be created.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="routedEventArgs">The event arguments</param>
        public void OkClicked(object sender, RoutedEventArgs routedEventArgs)
        {
            if (this.Validate())
            {
                this.DialogResult = true;
                this.Close();
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// This method is called by the Set accessor of each property.
        /// The CallerMemberName attribute that is applied to the optional propertyName
        /// parameter causes the property name of the caller to be substituted as an argument.
        /// </summary>
        /// <param name="propertyName">The name of the property that changed.</param>
        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Validates the settings for the recurring time entries.
        /// </summary>
        /// <returns><code>true</code> if the settings are valid and can be used to create recurring time entries, <code>false</code> otherwise.</returns>
        private bool Validate()
        {
            this.ValidationMessages.Clear();
            if (this.SelectedIssue == null)
            {
                this.ValidationMessages.Add("Es wurde kein Ticket ausgewählt.");
            }
            if (this.StartDate.Date > this.EndDate.Date)
            {
                this.ValidationMessages.Add("Der Starttag liegt nach dem Endtag");
            }
            if (this.StartTime > this.EndTime)
            {
                this.ValidationMessages.Add("Die angegebenen Buchungszeiten sind ungültig.");
            }
            this.NotifyPropertyChanged("ValidationMessagesString");
            return this.ValidationMessages.Count == 0;
        }

        #endregion
    }
}