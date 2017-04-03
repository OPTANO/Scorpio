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

namespace Scorpio.Outlook.Addin.Tests.LocalObjects
{
    using NUnit.Framework;

    using Scorpio.Outlook.AddIn.LocalObjects;

    /// <summary>
    /// Test for the issue info object
    /// </summary>
    [TestFixture]
    public class IssueInfoTests
    {
        /// <summary>
        /// Test to test the issue string
        /// </summary>
        [Test]
        public void TestIssueString()
        {
            // arrange
            var issueInfo = new IssueInfo() { Id = 4, Name = "Name", ProjectShortName = "ProjectShortName", ProjectId = 5, };

            // act
            var displayValue = issueInfo.IssueString;

            // assert
            Assert.That(displayValue, Is.EqualTo("#4"));
        }

        /// <summary>
        /// Method to test for a correct display value
        /// </summary>
        [Test]
        public void TestDisplayValue()
        {
            // arrange
            var issueInfo = new IssueInfo() { Id = 4, Name = "Name", ProjectShortName = "ProjectShortName", ProjectId = 5, };

            // act
            var displayValue = issueInfo.DisplayValue;

            // assert
            Assert.That(displayValue, Is.EqualTo("#4 - Name - [ProjectShortName]"));
        }
    }
}