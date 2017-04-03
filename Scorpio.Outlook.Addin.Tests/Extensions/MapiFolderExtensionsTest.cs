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

namespace Scorpio.Outlook.Addin.Tests.Extensions
{
    using System;

    using NUnit.Framework;

    using Scorpio.Outlook.AddIn.Extensions;

    /// <summary>
    /// Tests for the mapi folder extension
    /// </summary>
    [TestFixture]
    public class MapiFolderExtensionsTest
    {
        /// <summary>
        /// Method to test the filter string when start and end are included
        /// </summary>
        [Test]
        public void TestFilterStringIncludeStartIncludeEnd()
        {
            // arrange
            var startDate = new DateTime(2016, 7, 8);
            var endDate = new DateTime(2018, 9, 10);
            var includeStart = true;
            var includeEnd = true;
            var expectedFilterString = "[Start] <= '10.09.2018 00:00' AND [End] >= '08.07.2016 00:00'";

            // act
            var filterString = MapiFolderExtensions.GetFilterString(startDate, endDate, includeStart, includeEnd);

            // assert
            Assert.That(filterString, Is.EqualTo(expectedFilterString));
        }

        /// <summary>
        /// Method to test the filter string when start and end are included
        /// </summary>
        [Test]
        public void TestFilterStringIncludeStartNotIncludeEnd()
        {
            // arrange
            var startDate = new DateTime(2016, 7, 8);
            var endDate = new DateTime(2018, 9, 10);
            var includeStart = true;
            var includeEnd = false;
            var expectedFilterString = "[Start] < '10.09.2018 00:00' AND [End] >= '08.07.2016 00:00'";

            // act
            var filterString = MapiFolderExtensions.GetFilterString(startDate, endDate, includeStart, includeEnd);

            // assert
            Assert.That(filterString, Is.EqualTo(expectedFilterString));
        }

        /// <summary>
        /// Method to test the filter string when start and end are included
        /// </summary>
        [Test]
        public void TestFilterStringNotIncludeStartIncludeEnd()
        {
            // arrange
            var startDate = new DateTime(2016, 7, 8);
            var endDate = new DateTime(2018, 9, 10);
            var includeStart = false;
            var includeEnd = true;
            var expectedFilterString = "[Start] <= '10.09.2018 00:00' AND [End] > '08.07.2016 00:00'";

            // act
            var filterString = MapiFolderExtensions.GetFilterString(startDate, endDate, includeStart, includeEnd);

            // assert
            Assert.That(filterString, Is.EqualTo(expectedFilterString));
        }

        /// <summary>
        /// Method to test the filter string when start and end are included
        /// </summary>
        [Test]
        public void TestFilterStringNotIncludeStartNotIncludeEnd()
        {
            // arrange
            var startDate = new DateTime(2016, 7, 8);
            var endDate = new DateTime(2018, 9, 10);
            var includeStart = false;
            var includeEnd = false;
            var expectedFilterString = "[Start] < '10.09.2018 00:00' AND [End] > '08.07.2016 00:00'";

            // act
            var filterString = MapiFolderExtensions.GetFilterString(startDate, endDate, includeStart, includeEnd);

            // assert
            Assert.That(filterString, Is.EqualTo(expectedFilterString));
        }
    }
}