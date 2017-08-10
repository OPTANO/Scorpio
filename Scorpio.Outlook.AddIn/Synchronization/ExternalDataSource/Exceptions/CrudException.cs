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

namespace Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource.Exceptions
{
    using System;
    
    using Scorpio.Outlook.AddIn.LocalObjects;

    /// <summary>
    /// The exception thrown if there has been an error during the object creation in the external source
    /// </summary>
    public class CrudException : ScorpioException
    {
        /// <summary>
        /// Text to be displayed in the message
        /// </summary>
        private const string MessageText = "Error during CRUD operation on object";

        /// <summary>
        /// Initializes a new instance of the <see cref="CrudException"/> class.
        /// </summary>
        /// <param name="type">the type of operation performed</param>
        /// <param name="correspondingObject">
        /// The object to be created.
        /// </param>
        /// <param name="innerException">the inner exception</param>
        public CrudException(OperationType type, TimeEntryInfo correspondingObject, Exception innerException) : base(MessageText, innerException)
        {
            this.OperationType = type;
            this.CorrespondingObject = correspondingObject;
        }

        /// <summary>
        /// Gets the object to be created where the creation was not successful
        /// </summary>
        public TimeEntryInfo CorrespondingObject { get; private set; }

        /// <summary>
        /// Gets the operation type
        /// </summary>
        public OperationType OperationType { get; private set; }
        
    }
}