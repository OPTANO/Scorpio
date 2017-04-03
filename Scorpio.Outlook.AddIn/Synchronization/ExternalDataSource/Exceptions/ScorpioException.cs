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

    /// <summary>
    /// Base class for all exceptions in scorpio
    /// </summary>
    public class ScorpioException : Exception
    {
        /// <summary>
        /// The message text to be shown in the error
        /// </summary>
        private const string MessageText = "Error in SCORPIO";

        /// <summary>
        /// Initializes a new instance of the <see cref="ScorpioException"/> class.
        /// </summary>
        /// <param name="messageText">the message text of the exception, if none is set, the default one is used</param>
        /// <param name="baseException">
        /// The base exception.
        /// </param>
        public ScorpioException(string messageText, Exception baseException)
            : base(string.IsNullOrWhiteSpace(messageText) ? MessageText : messageText, baseException)
        {
        }

        /// <summary>
        /// Gets or sets the identifier number, can later be changed to an enum
        /// </summary>
        public int IdentifierNumber { get; set; }
    }
}