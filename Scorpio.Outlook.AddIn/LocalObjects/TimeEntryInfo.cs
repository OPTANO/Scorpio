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

namespace Scorpio.Outlook.AddIn.LocalObjects
{
    using System;

    /// <summary>
    /// The time entry info.
    /// </summary>
    [Serializable]
    public class TimeEntryInfo : AbstractInfoBase
    {
        /// <summary>
        /// Gets or sets the update time
        /// </summary>
        public DateTime UpdateTime { get; set; }

        /// <summary>
        /// Gets or sets the project info
        /// </summary>
        public ProjectInfo ProjectInfo { get; set; }

        /// <summary>
        /// Gets or sets the activity info
        /// </summary>
        public ActivityInfo ActivityInfo { get; set; }

        /// <summary>
        /// Gets or sets the issue info
        /// </summary>
        public IssueInfo IssueInfo { get; set; }

        /// <summary>
        /// Gets or sets the start date time
        /// </summary>
        public DateTime StartDateTime { get; set; }

        /// <summary>
        /// Gets or sets the end date time
        /// </summary>
        public DateTime EndDateTime { get; set; }
    }
}