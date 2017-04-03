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

namespace Scorpio.Outlook.Addin.Tests.Extensions
{
    using Microsoft.Office.Interop.Outlook;

    using Moq;

    using NUnit.Framework;

    using Scorpio.Outlook.AddIn.Extensions;

    /// <summary>
    /// Test fixture for the <see cref="AppointmentExtensions"/> class.
    /// </summary>
    [TestFixture]
    public class AppointmentExtensionsTest
    {
        #region Public Methods and Operators

        /// <summary>
        /// Tests the <see cref="AppointmentExtensions.AppendToBody"/> method.
        /// </summary>
        [Test]
        public void AppendToBody()
        {
            var appointmentMock = new Mock<AppointmentItem>();
            appointmentMock.SetupProperty(x => x.Body);

            var appointment = appointmentMock.Object;
            appointment.AppendToBody("foo");

            Assert.AreEqual("foo", appointment.Body);
        }

        /// <summary>
        /// Tests the <see cref="AppointmentExtensions.AppendToBody"/> method.
        /// </summary>
        [Test]
        public void AppendToBodyPreservesExistingText()
        {
            var appointmentMock = new Mock<AppointmentItem>();
            appointmentMock.SetupProperty(x => x.Body);

            var appointment = appointmentMock.Object;
            appointment.Body = "foo";

            appointment.AppendToBody("bar");

            Assert.AreEqual("foo\nbar", appointment.Body);
        }

        #endregion
    }
}