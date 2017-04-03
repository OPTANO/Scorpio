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

namespace Scorpio.Outlook.Addin.Tests.Synchronization.ExternalDataSource
{
    using NUnit.Framework;

    using Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource;

    /// <summary>
    /// Tests for the external data source factory
    /// </summary>
    [TestFixture]
    public class ExternalDataSourceFactoryTests
    {
        #region Constants

        /// <summary>
        /// The api key to use
        /// </summary>
        private const string ApiKey = "0a074de714f3ddee51d64cd8c8a0722846bf535d";

        /// <summary>
        /// The url to use
        /// </summary>
        private const string Url = @"https://services.orconomy.de/redmine";

        #endregion



        /// <summary>
        /// Test for getting the test instance
        /// </summary>
        [Test]
        public void GetTestInstance()
        {
            // arrange
            ExternalDataSourceFactory.UseTestManager = true;

            // act
            var instance = ExternalDataSourceFactory.GetRedmineMangerInstance(Url, ApiKey, 100);

            // assert
            Assert.That(instance, Is.Not.Null);
            Assert.That(instance.GetType(), Is.EqualTo(typeof(LocalListsExternalDataSourceTest)));
        }

        /// <summary>
        /// Test for getting the default instance
        /// </summary>
        [Test]
        public void GetDefaultInstance()
        {
            // arrange
            ExternalDataSourceFactory.UseTestManager = false;

            // act
            var instance = ExternalDataSourceFactory.GetRedmineMangerInstance(Url, ApiKey, 100);

            // assert
            Assert.That(instance, Is.Not.Null);
            Assert.That(instance.GetType(), Is.EqualTo(typeof(RedmineManagerInstance)));
        }
    }
}