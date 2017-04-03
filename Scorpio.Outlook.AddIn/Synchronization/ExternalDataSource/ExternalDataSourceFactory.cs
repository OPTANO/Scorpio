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

namespace Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource
{
    /// <summary>
    /// Factory for providing a redmine manager instance
    /// </summary>
    public class ExternalDataSourceFactory
    {
        /// <summary>
        /// The manager class to use
        /// </summary>
        private static IExternalSource manager;
        
        #region Public Methods and Operators

        /// <summary>
        /// Initializes a new instance of the <see cref="ExternalDataSourceFactory"/> class.
        /// </summary>
        /// <param name="address">
        /// The host address.
        /// </param>
        /// <param name="apiKey">
        /// The api key.
        /// </param>
        /// <param name="limitForNumberIssues">the limit to use for the number of issues to download</param>
        private ExternalDataSourceFactory(string address, string apiKey, int limitForNumberIssues)
        {
            if (UseTestManager)
            {
                manager = new LocalListsExternalDataSourceTest();
                manager.Limit = limitForNumberIssues;
            }
            else
            {
                manager = new RedmineManagerInstance(address, apiKey, limitForNumberIssues);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether a test manager should be used
        /// </summary>
        internal static bool UseTestManager { get; set; }

        /// <summary>
        /// Method to get a redmine manager instance
        /// </summary>
        /// <param name="address">the host address</param>
        /// <param name="apiKey">the api key</param>
        /// <param name="limitForNumber">the limit to use for the number of issues</param>
        /// <returns>the redmine manager</returns>
        public static IExternalSource GetRedmineMangerInstance(string address, string apiKey, int limitForNumber)
        {
            var factory = new ExternalDataSourceFactory(address, apiKey, limitForNumber);
            return manager;
            
        }

        #endregion
    }
}