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

namespace Scorpio.Outlook.AddIn.LocalObjects
{
    using System;

    /// <summary>
    /// Class that encapsulates information for displaying issue and project information on a time entry appointment.
    /// </summary>
    [Serializable]
    public class IssueInfo : AbstractInfoBase
    {
        #region Public properties
        
        /// <summary>
        /// Gets the issue id with the sharp sign in front, as a sting.
        /// </summary>
        public string IssueString
        {
            get
            {
                return "#" + this.Id;
            }
        }
        
        /// <summary>
        /// Gets the display name for the issue.
        /// </summary>
        public string DisplayValue
        {
            get
            {
                return string.Format("#{0} - {1} - [{2}]", this.Id, this.Name, this.ProjectShortName);
            }
        }

        /// <summary>
        /// Gets or sets the id of the corresponding project
        /// </summary>
        public int ProjectId { get; set; }

        /// <summary>
        /// Gets or sets the short name of the corresponding project
        /// </summary>
        public string ProjectShortName { get; set; }

        #endregion

        /// <summary>Returns a string that represents the current object.</summary>
        /// <returns>A string that represents the current object.</returns>
        /// <filterpriority>2</filterpriority>
        public override string ToString()
        {
            return string.Format("{0}, {3}, {1}, {2}", this.Id, this.ProjectId, this.ProjectShortName, this.Name);
        }
    }
}