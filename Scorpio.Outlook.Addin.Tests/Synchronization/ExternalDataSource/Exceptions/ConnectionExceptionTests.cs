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

namespace Scorpio.Outlook.Addin.Tests.Synchronization.ExternalDataSource.Exceptions
{
    using System;

    using NUnit.Framework;

    using Scorpio.Outlook.AddIn.Misc;
    using Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource.Exceptions;

    /// <summary>
    /// Tests for the exception
    /// </summary>
    [TestFixture]
    public class ConnectionExceptionTests : ScorpioExceptionTests
    {
        /// <summary>
        /// Method to test the exception
        /// </summary>
        [Test]
        public void TestConnectionException()
        {
            // arrange
            var inner = new Exception();

            // act
            var exception = new ConnectionException(inner);
            exception.PreviousAppointmentState = AppointmentState.Modified;

            // assert
            Assert.That(exception.InnerException, Is.EqualTo(inner));
            Assert.That(exception.PreviousAppointmentState, Is.EqualTo(AppointmentState.Modified));
        }
    }
}