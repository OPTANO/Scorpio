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
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows;

    using DevExpress.Xpf.Core;

    using Microsoft.Office.Interop.Outlook;

    using Scorpio.Outlook.AddIn.Extensions;

    /// <summary>
    /// Interaction logic for SaveDialog.xaml
    /// </summary>
    public partial class SaveDialog : DXWindow
    {
        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SaveDialog"/> class.
        /// </summary>
        /// <param name="items">The elements which are modified</param>
        public SaveDialog(IEnumerable<AppointmentItem> items)
        {
            this.DataContext = this;
            this.DeletedItems =
                items.Where(i => i.IsDeletedSet())
                    .Select(i => new TimeEntryDetails() { End = i.End, Start = i.Start, Subject = i.Subject, Location = i.Location })
                    .ToList();
            this.InitializeComponent();
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Gets or sets the elements that correspond to deleted time entries.
        /// </summary>
        public List<TimeEntryDetails> DeletedItems { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// User clicked the cancel button
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The arguments</param>
        private void CancelClicked(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        /// <summary>
        /// User clicked the ok button
        /// </summary>
        /// <param name="sender">The sender</param>
        /// <param name="e">The arguments</param>
        private void OkClicked(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        #endregion
    }
}