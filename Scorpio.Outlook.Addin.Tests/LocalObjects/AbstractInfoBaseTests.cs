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
    /// Tests for the abstract base info elements
    /// </summary>
    [TestFixture]
    public class AbstractInfoBaseTests
    {
        /// <summary>
        /// Method to test if two equal objects are seen as equal
        /// </summary>
        [Test]
        public void TestEquality()
        {
            // arrange
            var projectInfoOne = new ProjectInfo() { Id = 5, Name = "Test" };
            var projectInfoTwo = new ProjectInfo() { Id = 5, Name = "Test" };

            // act
            var areEqual = object.Equals(projectInfoTwo, projectInfoOne);

            // assert
            Assert.That(areEqual, Is.True);
        }

        /// <summary>
        /// Method to test if two equal objects are not seen as equal if they have another name
        /// </summary>
        [Test]
        public void TestEqualityOtherName()
        {
            // arrange
            var projectInfoOne = new ProjectInfo() { Id = 5, Name = "Test" };
            var projectInfoTwo = new ProjectInfo() { Id = 5, Name = "Test2" };

            // act
            var areEqual = object.Equals(projectInfoTwo, projectInfoOne);

            // assert
            Assert.That(areEqual, Is.False);
        }

        /// <summary>
        /// Method to test if two equal objects are seen as not equal if they have anotherid
        /// </summary>
        [Test]
        public void TestEqualityOtherId()
        {
            // arrange
            var projectInfoOne = new ProjectInfo() { Id = 7, Name = "Test" };
            var projectInfoTwo = new ProjectInfo() { Id = 5, Name = "Test" };

            // act
            var areEqual = object.Equals(projectInfoTwo, projectInfoOne);
            var hashOne = projectInfoOne.GetHashCode();
            var hashTwo = projectInfoTwo.GetHashCode();

            // assert
            Assert.That(areEqual, Is.False);
            Assert.That(hashOne, Is.Not.EqualTo(hashTwo));
        }

        /// <summary>
        /// Method to test if two equal objects having the same values but another type are equal
        /// </summary>
        [Test]
        public void TestEqualityOtherType()
        {
            // arrange
            AbstractInfoBase projectInfo = new ProjectInfo() { Id = 7, Name = "Test" };
            AbstractInfoBase userInfo = new UserInfo() { Id = 5, Name = "Test" };

            // act
            var areEqual = object.Equals(userInfo, projectInfo);

            // assert
            Assert.That(areEqual, Is.False);
        }


    }
}