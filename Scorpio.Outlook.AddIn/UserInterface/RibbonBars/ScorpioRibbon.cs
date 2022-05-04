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

namespace Scorpio.Outlook.AddIn.UserInterface.RibbonBars
{
    using System;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;

    using log4net;

    /// <summary>
    /// Ribbon menu for SCORPIO.
    /// </summary>
    [ComVisible(true)]
    public partial class ScorpioRibbon : Office.IRibbonExtensibility
    {
        #region Static Fields

        /// <summary>
        /// The logger.
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(ScorpioRibbon));

        #endregion

        #region Fields

        /// <summary>
        /// The ribbon control.
        /// </summary>
        private Office.IRibbonUI ribbon;

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Method that returns the custom ui for the ribbon extension.
        /// </summary>
        /// <param name="ribbonID">The ribbon id</param>
        /// <returns>The ribbon ui as a string.</returns>
        public string GetCustomUI(string ribbonID)
        {
            if ("Microsoft.Outlook.Explorer".Equals(ribbonID))
            {
                return GetResourceText("Scorpio.Outlook.AddIn.UserInterface.RibbonBars.ScorpioRibbonExplorer.xml");
            }
            if ("Microsoft.Outlook.Appointment".Equals(ribbonID))
            {
                return GetResourceText("Scorpio.Outlook.AddIn.UserInterface.RibbonBars.ScorpioRibbonAppointment.xml");
                ;
            }
            return null;
        }

        /// <summary>
        /// Called when the ribbon is loaded.
        /// </summary>
        /// <param name="ribbonUI">The ribbon ui for the ribbon.</param>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Gets a resource as text.
        /// </summary>
        /// <param name="resourceName">The name of the resource.</param>
        /// <returns>The resource as text.</returns>
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}