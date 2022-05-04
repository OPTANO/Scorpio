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
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using System.Windows;
    using System.Windows.Input;

    using DevExpress.Mvvm.POCO;

    using log4net;

    using Scorpio.Outlook.AddIn.Cache;
    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Misc;
    using Scorpio.Outlook.AddIn.Properties;
    using Scorpio.Outlook.AddIn.UserInterface.Helper;

    /// <summary>
    /// Interaction logic for SettingsDialog.xaml
    /// </summary>
    public partial class SettingsEditorDialog : INotifyPropertyChanged
    {
        #region Fields

        /// <summary>
        /// Logging for error logging
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(SettingsEditorDialog));

        /// <summary>
        /// Private field for the redmine API key.
        /// </summary>
        private string _redmineApiKey;

        /// <summary>
        /// Private field for the redmine URL.
        /// </summary>
        private string _redmineUrl;

        /// <summary>
        /// Private field for the command which removes an issue from the list of favorite issues.
        /// </summary>
        private ICommand _removeFavoriteCommand;

        /// <summary>
        /// The selected issue.
        /// </summary>
        private IssueInfo _selectedIssue;

        /// <summary>
        /// Backing field for the refresh time
        /// </summary>
        private double _refreshTime;

        /// <summary>
        /// The number of issues to load
        /// </summary>
        private int _numberIssues;

        /// <summary>
        /// The number of issues to show at most in the last used issue area
        /// </summary>
        private int _numberIssuesLastUsed;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SettingsEditorDialog"/> class.
        /// </summary>
        public SettingsEditorDialog()
        {
            var stopWatch = Stopwatch.StartNew();
            stopWatch.Start();

            this.InitializeComponent();
            if (!this.IsInDesignMode())
            {
                Debug.WriteLine($"InitializeComponent after {stopWatch.ElapsedMilliseconds}ms");


                this.RedmineApiKey = Settings.Default.RedmineApiKey;
                this.RedmineUrl = Settings.Default.RedmineURL;
                try
                {
                    this.FavoriteIssues = new ObservableCollection<IssueInfo>(Globals.ThisAddIn.Synchronizer.FavoriteIssues);
                }
                catch (Exception e)
                {
                    this.FavoriteIssues = new ObservableCollection<IssueInfo>();
                    Log.Error("Error while setting the favorite issues in the settings dialog", e);
                }
                Debug.WriteLine($"FacoriteIssues after {stopWatch.ElapsedMilliseconds}ms");

                try
                {
                    this.AvailableIssues = new ObservableCollection<IssueInfo>(Globals.ThisAddIn.Synchronizer.AllIssues.Values);
                }
                catch (Exception e)
                {
                    this.AvailableIssues = new ObservableCollection<IssueInfo>();
                    Log.Error("Error while setting the overall issue list in the settings dialog", e);
                }
                Debug.WriteLine($"AvailableIssues after {stopWatch.ElapsedMilliseconds}ms");

                this.RefreshTime = Settings.Default.RefreshTime;
                this.NumberIssues = Settings.Default.LimitForIssueNumber;
                this.NumberIssuesLastUsed = Settings.Default.NumberLastUsedIssues;
                Debug.WriteLine($"Finished ctor after {stopWatch.ElapsedMilliseconds}ms");
                stopWatch.Stop();
                this.DataContext = this;
            }
         
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
        /// Gets or sets the issues available for using as favorite issues.
        /// </summary>
        public ObservableCollection<IssueInfo> AvailableIssues { get; set; }

        /// <summary>
        /// Gets or sets the issues that are marked as favorite issues.
        /// </summary>
        public ObservableCollection<IssueInfo> FavoriteIssues { get; set; }

        /// <summary>
        /// Gets or sets the refresh time
        /// </summary>
        public double RefreshTime
        {
            get
            {
                return this._refreshTime;
            }
            set
            {
                if (!this._refreshTime.Equals(value))
                {
                    this._refreshTime = value;
                    this.NotifyPropertyChanged();
                }
            }
        }
        
        /// <summary>
        /// Gets or sets the maximum amount of issues shown in the last used area
        /// </summary>
        public int NumberIssuesLastUsed
        {
            get
            {
                return this._numberIssuesLastUsed;
            }
            set
            {
                if (!this._numberIssuesLastUsed.Equals(value))
                {
                    this._numberIssuesLastUsed = Math.Max(0, value);
                    Settings.Default.NumberLastUsedIssues = this._numberIssuesLastUsed;
                    Settings.Default.Save();
                    this.NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Gets or sets the refresh time
        /// </summary>
        public int NumberIssues
        {
            get
            {
                return this._numberIssues;
            }
            set
            {
                if (!this._numberIssues.Equals(value))
                {
                    this._numberIssues = Math.Max(0, value);
                    Settings.Default.LimitForIssueNumber = this._numberIssues;
                    Settings.Default.Save();
                    this.NotifyPropertyChanged();
                }
            }
        }
        

        /// <summary>
        /// Gets or sets the Api key for Redmine.
        /// </summary>
        public string RedmineApiKey
        {
            get
            {
                return this._redmineApiKey;
            }
            set
            {
                if (this._redmineApiKey != value)
                {
                    this._redmineApiKey = value;
                    this.NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Gets or sets the url of Redmine.
        /// </summary>
        public string RedmineUrl
        {
            get
            {
                return this._redmineUrl;
            }
            set
            {
                if (value != this._redmineUrl)
                {
                    this._redmineUrl = value;
                    this.NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Gets the command which removes the specified issue from the list of favorite issues.
        /// </summary>
        public ICommand RemoveFavoriteCommand
        {
            get
            {
                return this._removeFavoriteCommand
                       ?? (this._removeFavoriteCommand = new RelayCommand<IssueInfo>(i => this.FavoriteIssues.Remove(i), i => true));
            }
        }

        /// <summary>
        /// Gets or sets the issue which is selected to be added as a new favorite issue.
        /// </summary>
        public IssueInfo SelectedIssue
        {
            get
            {
                return this._selectedIssue;
            }
            set
            {
                if (this._selectedIssue != value)
                {
                    this._selectedIssue = value;
                    this.NotifyPropertyChanged();
                }
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Method that is called when the user wants to add an issue to their favorite issues.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The event arguments</param>
        private void AddFavorite_OnClick(object sender, RoutedEventArgs e)
        {
            if (this.SelectedIssue != null)
            {
                if (!this.FavoriteIssues.Contains(this.SelectedIssue))
                {
                    this.FavoriteIssues.Add(this.SelectedIssue);
                }
                this.SelectedIssue = null;
            }
        }

        /// <summary>
        /// Closes the dialog, indicating that the changes should not be applied.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The event arguments</param>
        private void Cancel_OnClick(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        /// <summary>
        /// Clears the issue cache. Called when the user clicks the corresponding button in the dialog.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The event arguments</param>
        private void ClearIssueCache_OnClick(object sender, RoutedEventArgs e)
        {
            Settings.Default.LastIssueSyncDate = new DateTime(1900, 01, 01);
            Settings.Default.Save();
            LocalCache.DeleteEntry(LocalCache.KnownProjects);
            LocalCache.DeleteEntry(LocalCache.KnownIssues);
        }

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
        /// Called when the user clicks the OK button. Saves all changed settings to the settings file and closes the dialog.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The event arguments</param>
        private void Ok_OnClick(object sender, RoutedEventArgs e)
        {
            Settings.Default.RedmineApiKey = this.RedmineApiKey;
            Settings.Default.RedmineURL = this.RedmineUrl;
            Settings.Default.LimitForIssueNumber = Math.Max(1, this.NumberIssues);
            Settings.Default.RefreshTime = Math.Max(0, this.RefreshTime);

            Settings.Default.Save();

            Globals.ThisAddIn.Synchronizer.UpdateFavoriteIssues(this.FavoriteIssues.ToList());

            this.DialogResult = true;
            this.Close();
        }

        #endregion
    }
}