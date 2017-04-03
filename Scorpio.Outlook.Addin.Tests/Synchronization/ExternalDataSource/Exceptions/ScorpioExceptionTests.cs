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

    using Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource.Exceptions;

    /// <summary>
    /// Tests for the exception
    /// </summary>
    [TestFixture]
    public class ScorpioExceptionTests
    {
        /// <summary>
        /// Test for the empty case (no message text, no inner exception)
        /// </summary>
        [Test]
        public void TestMessageNoTextAndException()
        {
            // arrange
            
            // act
            var exception = new ScorpioException(null, null);

            // assert
            Assert.That(exception.Message, Contains.Substring("SCORPIO"));
            Assert.That(exception.InnerException, Is.Null);
        }

        /// <summary>
        /// Test for the case of no text and an inner exception
        /// </summary>
        [Test]
        public void TestMessageInnerException()
        {
            // arrange
            var innerException = new Exception();

            // act
            var exception = new ScorpioException(null, innerException);

            // assert
            Assert.That(exception.Message, Contains.Substring("SCORPIO"));
            Assert.That(exception.InnerException, Is.EqualTo(innerException));
        }

        /// <summary>
        /// Test for the case of text and no inner exception
        /// </summary>
        [Test]
        public void TestMessageTextAndException()
        {
            // arrange
            var text = "Hallo";
            var innerException = new Exception();

            // act
            var exception = new ScorpioException(text, innerException);
            exception.IdentifierNumber = 5;

            // assert
            Assert.That(exception.InnerException, Is.EqualTo(innerException));
            Assert.That(exception.Message, Is.EqualTo(text));
            Assert.That(exception.IdentifierNumber, Is.EqualTo(5));
        }

        /// <summary>
        /// Test for the case of text and no inner exception
        /// </summary>
        [Test]
        public void TestMessageText()
        {
            // arrange
            var text = "Hallo";

            // act
            var exception = new ScorpioException(text, null);

            // assert
            Assert.That(exception.InnerException, Is.Null);
            Assert.That(exception.Message, Is.EqualTo(text));
        }
    }
}