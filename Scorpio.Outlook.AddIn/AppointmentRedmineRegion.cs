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

using Office = Microsoft.Office.Core;

namespace Scorpio.Outlook.AddIn
{
    using System;
    using System.Collections.ObjectModel;
    using System.Diagnostics;
    using System.Linq;
    using System.Windows.Forms;

    using DevExpress.Data.Filtering;
    using DevExpress.Mvvm.POCO;
    using DevExpress.Utils.Win;
    using DevExpress.XtraEditors;
    using DevExpress.XtraLayout;
    using DevExpress.XtraLayout.Utils;

    using log4net;

    using Microsoft.Office.Interop.Outlook;
    using Microsoft.Office.Tools.Outlook;

    using Scorpio.Outlook.AddIn.Extensions;
    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Misc;

    using Exception = System.Exception;

    /// <summary>
    /// Formregion that adds redmine functionalilty to the appointment window of outlook.
    /// </summary>
    public partial class AppointmentRedmineRegion
    {
        #region Static Fields

        /// <summary>
        /// The logger.
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(AppointmentRedmineRegion));

        #endregion

        /// <summary>
        /// The list containing the issues
        /// </summary>
        private ObservableCollection<IssueInfo> issueList;


        #region Public Methods and Operators

        /// <summary>
        /// Method that allows to set the focus to the issue selector.
        /// </summary>
        public void FocusIssueSelection()
        {
            this.issueSelector.Focus();
            this.issueSelector.ShowPopup();
        }

        #endregion

        #region Methods

        /// <summary>
        /// Occurs when the form region is closed.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The event arguments.</param>
        private void AppointmentRedmineRegion_FormRegionClosed(object sender, EventArgs e)
        {
        }
        
        /// <summary>
        /// Occurs before the form region is displayed.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The event arguments.</param>
        private void AppointmentRedmineRegion_FormRegionShowing(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Synchronizer != null && Globals.ThisAddIn.Synchronizer.AllIssues != null
                && Globals.ThisAddIn.Synchronizer.LastUsedIssues != null && Globals.ThisAddIn.Synchronizer.FavoriteIssues != null)
            {
                this.issueSelector.Popup += this.IssueSelectorOnPopup;

                this.issueList = new ObservableCollection<IssueInfo>(Globals.ThisAddIn.Synchronizer.AllIssues.Values);
                this.issueProjectInfoBindingSource.DataSource = this.issueList;
                this.lastUsedIssuesBindingSource.DataSource = Globals.ThisAddIn.Synchronizer.LastUsedIssues; 
                this.favoriteIssuesBindingSource.DataSource = Globals.ThisAddIn.Synchronizer.FavoriteIssues;
                this.UpdateUi();
            }
            else
            {
                Log.ErrorFormat("Opening Appointment Region failed. Redmine Synchronizer or its data is not initialized yet. ");
                this.Visible = false;
            }
        }

        /// <summary>
        /// Is invoked after the edit value changed
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The event args</param>
        private void IssueSelector_EditValueChanged(object sender, EventArgs e)
        {
            var issueId = this.issueSelector.EditValue;
            if (issueId == null)
            {
                return;
            }
            int id;
            try
            {
                id = Convert.ToInt32(issueId);
            }
            catch (Exception ex)
            {
                Log.Error("Could not parse IssueId from Issue Selector", ex);
                return;
            }

            this.SetIssue(id);
        }

        /// <summary>
        /// When the popup opens, we have to attach a Keyeventhandler to the editor of the find panel.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="eventArgs">The event arguments</param>
        private void IssueSelectorOnPopup(object sender, EventArgs eventArgs)
        {
            try
            {
                // There does not seem to be a better way to get to the required data of the TextEdit that contains the filter string. 
                // See https://www.devexpress.com/Support/Center/Question/Details/Q362889
                var editor = (sender as DevExpress.Utils.Win.IPopupControl).PopupWindow.Controls[2].Controls[0].Controls[7] as TextEdit;
                editor.KeyDown += new KeyEventHandler(this.PopupFindBoxKeyDown);
                editor.EditValueChanged += this.TicketSearchEditorOnEditValueChanged;
            }
            catch (Exception e)
            {
                Log.Error(
                    "Could not attach KeyEventHandler to the find control of the popup. "
                    + "Selection by pressing Enter in the find control will not work.",
                    e);
            }
        }

        /// <summary>
        /// Handler method for the case that a link was clicked.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The event arguments.</param>
        private void LnkProject_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var startInfo = new ProcessStartInfo(e.Link.LinkData.ToString());
            Process.Start(startInfo);
        }

        /// <summary>
        /// Is invoked when the selected value of the last used issue lst changed
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The event args</param>
        private void LstFavorite_SelectedValueChanged(object sender, EventArgs e)
        {
            if (this.lstFavorite.SelectedIndex < 0)
            {
                return;
            }
            var id = (int)this.lstFavorite.SelectedValue;
            this.SetIssue(id);
            this.lstFavorite.SelectedIndex = -1;
        }

        /// <summary>
        /// Is invoked when the selected value of the last used issue lst changed
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The event args</param>
        private void LstLastUsed_SelectedValueChanged(object sender, EventArgs e)
        {
            if (this.lstLastUsed.SelectedIndex < 0)
            {
                return;
            }
            var idx = (int)this.lstLastUsed.SelectedValue;
            this.SetIssue(idx);
            this.lstLastUsed.SelectedIndex = -1;
        }

        /// <summary>
        /// Handles keypresses in the text editor searchfield of the SearchLookupEditView. If enter is pressed there, 
        /// we take the first element that is displayed in the grid, and set it as the selected value.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The event arguments</param>
        private void PopupFindBoxKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.issueSelector.Properties.View.FocusedRowHandle = 0;
                this.issueSelector.ClosePopup();
            }
        }

        /// <summary>
        /// Sets the issue id of the Outlook appointment which is edited in the editor.
        /// </summary>
        /// <param name="issueId">The issue id.</param>
        private void SetIssue(int issueId)
        {
            var appointment = this.OutlookItem as AppointmentItem;
            if (appointment != null)
            {
                Globals.ThisAddIn.Synchronizer.UpdateAppointmentIssue(appointment, issueId);
                this.UpdateUi();
                this.lstLastUsed.Refresh();
            }
        }

        /// <summary>
        /// Method that is executed everytime the editvalue of the search field in the searchlookupedit control for finding an issue is changed.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="eventArgs">The event arguments</param>
        private void TicketSearchEditorOnEditValueChanged(object sender, EventArgs eventArgs)
        {
            var editor = sender as TextEdit;
            if (editor == null)
            {
                return;
            }

            var filterText = editor.Text;
            // reload the issue, if needed
            var reloaded = this.ReloadIssueIfNeeded(filterText);
            if (reloaded && !filterText.StartsWith("#"))
            {
                filterText = string.Format("#{0}", filterText);
            }

            // set filter
            if (filterText.StartsWith("#"))
            {
                var filter = new BinaryOperator(new OperandProperty("IssueString"), new ConstantValue(filterText), BinaryOperatorType.Equal);
                this.issueSelector.Properties.View.ActiveFilterCriteria = filter;
            }
            else
            {
                this.issueSelector.Properties.View.ActiveFilterCriteria = null;
            }
        }

        /// <summary>
        /// Method to reload the issue infos for the id contained in teh filter text
        /// </summary>
        /// <param name="filterText">the filter text</param>
        /// <returns>if a reload was done</returns>
        private bool ReloadIssueIfNeeded(string filterText)
        {
            var textToUseForSearch = filterText.GetStringToUseForUnknownIssueSearch();
            if (textToUseForSearch == null)
            {
                // nothing to here
                return false;
            }
            var synchronizer = Globals.ThisAddIn.Synchronizer;
            int issueId;
            if (int.TryParse(textToUseForSearch, out issueId))
            {
                var issueAlreadyKnown = synchronizer.AllIssues.ContainsKey(issueId);
                if (!issueAlreadyKnown)
                {
                    // try to get the infos
                    var issueInfoAndNewIssueList = synchronizer.ReloadIssueById(issueId);
                    foreach (var issueInfo in issueInfoAndNewIssueList.Item2)
                    {
                        this.issueList.Add(issueInfo);
                    }

                    return issueInfoAndNewIssueList.Item2.Any();

                }
            }
            return false;
        }

        /// <summary>
        /// Updates the Ui with the redmine appointment information
        /// </summary>
        private void UpdateUi()
        {
            var appointment = this.OutlookItem as AppointmentItem;
            if (appointment != null)
            {
                var projectIdField = appointment.UserProperties.Find(Constants.FieldRedmineProjectId);
                var issueId = appointment.GetIssueId();

                if (projectIdField != null)
                {
                    var info = Globals.ThisAddIn.Synchronizer.BuildProjectInformation(Convert.ToInt32(projectIdField.Value));
                    this.lnkProject.Text = info.Item2;
                    this.lnkProject.Links.Remove(this.lnkProject.Links[0]);
                    this.lnkProject.Links.Add(0, this.lnkProject.Text.Length, info.Item1);
                    this.lnkProject.Enabled = true;
                }
                else
                {
                    this.lnkProject.Text = "-";
                    this.lnkProject.Enabled = false;
                }

                if (issueId.HasValue)
                {
                    var info = Globals.ThisAddIn.Synchronizer.BuildIssueInformation(issueId.Value);
                    this.lnkIssue.Text = info.Item2;
                    this.lnkIssue.Links.Remove(this.lnkIssue.Links[0]);
                    this.lnkIssue.Links.Add(0, this.lnkIssue.Text.Length, info.Item1);
                    this.lnkIssue.Enabled = true;

                    // set the current issue
                    this.issueSelector.EditValue = issueId.Value;
                }
                else
                {
                    this.lnkIssue.Text = "-";
                    this.lnkIssue.Enabled = false;
                }
            }
        }

        #endregion

        /// <summary>
        /// The factory.
        /// </summary>
        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Appointment)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("Scorpio.Outlook.AddIn.AppointmentRedmineRegion")]
        public partial class AppointmentRedmineRegionFactory
        {
            #region Methods

            /// <summary>
            /// Method that is called when the redmine form region for appointments is about to be initialized.
            /// </summary>
            /// <param name="sender">The sender.</param>
            /// <param name="e">The event parameters.</param>
            private void AppointmentRedmineRegionFactory_FormRegionInitializing(object sender, FormRegionInitializingEventArgs e)
            {
                // Occurs before the form region is initialized.
                // To prevent the form region from appearing, set e.Cancel to true.
                // Use e.OutlookItem to get a reference to the current Outlook item.
                var item = e.OutlookItem as AppointmentItem;

                var parent = item?.Parent as MAPIFolder;

                if (parent != null)
                {
                    if (parent.EntryID.Equals(Globals.ThisAddIn.RedmineCalendar.EntryID))
                    {
                        // we do not have to cancel, because the item is being created in the redmine calendar.
                        return;
                    }
                }

                // The item is not in the redmine calendar, thus cancel showing the form region.
                e.Cancel = true;
            }

            #endregion
        }

        /// <summary>
        /// Method to search in the last used issues list.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void searchLastIssues_TextChanged(object sender, EventArgs e)
        {
            this.lastUsedIssuesBindingSource.DataSource = Globals.ThisAddIn.Synchronizer.LastUsedIssues.Where(i => i.DisplayValue.ContainsAllWords(this.searchLastIssues.Text));
        }

        /// <summary>
        /// Method handles Popup Event from the isse Selector to hide the clear Button.
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The event args</param>
        private void issueSelector_Popup(object sender, EventArgs e)
        {
            // Simple:
            // this.issueSelector.Properties.View.OptionsFind.ShowClearButton = false;
            // Would be too easy....
            // Offical solution to hide a Button: 
            // https://www.devexpress.com/Support/Center/Question/Details/Q393620/gridlookupedit-popupfind-hide-find-button
            // o.O
            LayoutControl lc = (sender as IPopupControl).PopupWindow.Controls[2].Controls[0] as LayoutControl;
            ((lc.Items[0] as LayoutControlGroup).Items[2] as LayoutControlGroup).Items[0].Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            
        }
    }
}