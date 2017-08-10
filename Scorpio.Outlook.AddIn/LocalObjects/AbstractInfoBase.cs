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
    /// The base class for all info objects
    /// </summary>
    [Serializable]
    public abstract class AbstractInfoBase
    {
        #region Public Properties

        /// <summary>
        /// Gets or sets the id of the object
        /// </summary>
        public int? Id { get; set; }

        /// <summary>
        /// Gets or sets the name of the object
        /// </summary>
        public string Name { get; set; }

        #endregion

        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>the hash code</returns>
        public override int GetHashCode()
        {
            return this.Id.GetValueOrDefault(-1).GetHashCode() ^ this.Name.GetHashCode();
        }

        /// <summary>
        /// The equals method
        /// </summary>
        /// <param name="obj">the object to compare to</param>
        /// <returns>if the objects are equal</returns>
        public override bool Equals(object obj)
        {
            var other = obj as AbstractInfoBase;
            var otherType = obj.GetType();
            var thisType = this.GetType();

            return other != null && object.Equals(this.Id, other.Id) && object.Equals(this.Name, other.Name) && object.Equals(otherType, thisType);
        }
    }
}