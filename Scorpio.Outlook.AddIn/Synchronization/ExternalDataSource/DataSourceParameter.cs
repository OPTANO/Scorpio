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
    using System;

    /// <summary>
    /// The data source parameter for queries
    /// </summary>
    public class DataSourceParameter
    {
        #region Public Properties

        /// <summary>
        /// Gets or sets the issue id
        /// </summary>
        public int? IssueId { get; set; }

        /// <summary>
        /// Gets or sets the maximum amount of items to retrieve
        /// </summary>
        public int? Limit { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the set limit should be used
        /// </summary>
        public bool? UseLimit { get; set; }

        /// <summary>
        /// Gets or sets the project id to filter for
        /// </summary>
        public int? ProjectId { get; set; }

        /// <summary>
        /// Gets or sets the start and end date and time of the spent on time range
        /// </summary>
        public Tuple<DateTime, DateTime> SpentDateTimeTuple { get; set; }

        /// <summary>
        /// Gets or sets the status id, -1 is interpreted as all
        /// </summary>
        public int? StatusId { get; set; }
        
        /// <summary>
        /// Gets or sets the start date time for the update time range
        /// </summary>
        public DateTime? UpdateStartDateTime { get; set; }

        /// <summary>
        /// Gets or sets the user id, -1 means me
        /// </summary>
        public int? UserId { get; set; }

        #endregion
    }
}