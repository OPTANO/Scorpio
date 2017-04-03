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

namespace Scorpio.Outlook.AddIn.UserInterface.RibbonBars
{
    using System.Drawing;

    using Office = Microsoft.Office.Core;

    /// <summary>
    /// Partial class for the scorpio ribbon
    /// </summary>
    public partial class ScorpioRibbon
    {
        #region Public Methods and Operators

        /// <summary>
        /// Gets an image for a ribbon control.
        /// </summary>
        /// <param name="control">The control for which to get the image.</param>
        /// <returns>The image for the control.</returns>
        public Bitmap GetImageForAppointmentRibbon(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "selectTicket":
                    return Properties.Resources.magnifier;
            }
            return null;
        }

        /// <summary>
        /// When called, looks for the appointment redmine region in the current inspector and tries to focus the issue selector.
        /// </summary>
        /// <param name="control">The ribbon control.</param>
        public void OnSelectTicket(Office.IRibbonControl control)
        {
            // Try to get the current inspector
            var inspector = Globals.ThisAddIn.Application.ActiveInspector();

            // If it is null, there is no appointment inspector open
            if (inspector == null)
            {
                return;
            }

            // If the appointment redmine region is null, we have an inspector that does not show the region
            var formRegions = Globals.FormRegions[inspector];
            formRegions.AppointmentRedmineRegion?.FocusIssueSelection();
        }

        #endregion
    }
}