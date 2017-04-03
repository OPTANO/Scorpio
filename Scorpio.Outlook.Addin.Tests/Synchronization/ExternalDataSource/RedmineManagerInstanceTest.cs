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
    using System;
    using System.Collections.Specialized;
    using System.Linq;

    using NUnit.Framework;

    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource;

    /// <summary>
    /// Tests for the redmine manager instance
    /// </summary>
    [TestFixture]
    public class RedmineManagerInstanceTest
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

        #region Public Methods and Operators

        /// <summary>
        /// Method to test the download of the activities
        /// </summary>
        [Test]
        public void TestDownloadActivityInfos()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameters = new DataSourceParameter();

            // act
            var activities = manager.GetTotalActivityInfoList(parameters);

            // assert
            Assert.That(activities, Is.Not.Null);
            Assert.That(activities.Count, Is.GreaterThan(0));
        }

        /// <summary>
        /// Method to test the download of the projects
        /// </summary>
        [Test]
        public void TestDownloadProjectInfos()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameters = new DataSourceParameter();

            // act
            var projects = manager.GetTotalProjectList(parameters);

            // assert
            Assert.That(projects, Is.Not.Null);
            Assert.That(projects.Count, Is.GreaterThan(0));
        }

        /// <summary>
        /// Method to test the download of the issues with a limit set as parameter
        /// </summary>
        [Test]
        public void TestDownloadIssueInfosWithLimit()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameters = new DataSourceParameter() { Limit = 1 };
            
            // act
            var issues = manager.GetIssueInfoList(parameters);

            // assert
            Assert.That(issues, Is.Not.Null);
            Assert.That(issues.Count, Is.GreaterThan(0));
            Assert.That(issues.Count, Is.LessThanOrEqualTo(1));
        }

        /// <summary>
        /// Method to test the download of the issues with a limit set as parameter
        /// </summary>
        [Test]
        public void TestDownloadIssueInfosWithLimitViaBoolLimit()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey, 1);
            var parameters = new DataSourceParameter() { UseLimit = true };

            // act
            var issues = manager.GetIssueInfoList(parameters);

            // assert
            Assert.That(issues, Is.Not.Null);
            Assert.That(issues.Count, Is.GreaterThan(0));
            Assert.That(issues.Count, Is.LessThanOrEqualTo(1));
        }

        /// <summary>
        /// Method to test the download of the issues with a filter set to an existing project id
        /// </summary>
        [Test]
        public void TestDownloadIssueInfosWithProjectFilterForExistingProject()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameters = new DataSourceParameter() { ProjectId = 333 };

            // act
            var issues = manager.GetTotalIssueInfoList(parameters);

            // assert
            Assert.That(issues, Is.Not.Null);
            Assert.That(issues.Count, Is.GreaterThan(0));
        }

        /// <summary>
        /// Method to test the download of the issues
        /// </summary>
        [Test]
        public void TestDownloadIssueInfos()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameters = new DataSourceParameter();

            // act
            var issues = manager.GetTotalIssueInfoList(parameters);

            // assert
            Assert.That(issues, Is.Not.Null);
            Assert.That(issues.Count, Is.GreaterThan(0));
        }

        /// <summary>
        /// Method to test the download of the time entries
        /// </summary>
        [Test]
        public void TestDownloadTimeEntryInfos()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameters = new DataSourceParameter();

            // act
            var timeEntries = manager.GetTotalTimeEntryInfoList(parameters);

            // assert
            Assert.That(timeEntries, Is.Not.Null);
            Assert.That(timeEntries.Count, Is.GreaterThan(0));
        }

        /// <summary>
        /// Test for the conversion of the empty parameter list
        /// </summary>
        [Test]
        public void TestParameterConversionNoParameterGiven()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameter = new DataSourceParameter();
            
            // act
            var queryParameter = manager.GetParametersForQuery(parameter);
            
            // assert
            Assert.That(queryParameter, Is.Not.Null);
            Assert.That(queryParameter.Count, Is.EqualTo(0));
        }

        /// <summary>
        /// Test for the conversion of the parameter containing only an user id
        /// </summary>
        /// <param name="userIdToSet">the id to set in the parameter</param>
        /// <param name="expectedValue">the value expected in the result</param>
        [TestCase(1, "1")]
        [TestCase(-1, "me")]
        public void TestParameterConversionNoParameterButUserIdSet(int userIdToSet, string expectedValue)
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameter = new DataSourceParameter() { UserId = userIdToSet };

            // act
            var queryParameter = manager.GetParametersForQuery(parameter);

            // assert
            Assert.That(queryParameter, Is.Not.Null);
            Assert.That(queryParameter.Count, Is.EqualTo(1));
            Assert.That(queryParameter["user_id"], Is.EqualTo(expectedValue));
        }

        /// <summary>
        /// Test for the conversion of the parameter containing only an user id
        /// </summary>
        /// <param name="statusIdToSet">the id to set in the parameter</param>
        /// <param name="expectedValue">the value expected in the result</param>
        [TestCase(13, "13")]
        [TestCase(-1, "*")]
        public void TestParameterConversionNoParameterButStatusIdSet(int statusIdToSet, string expectedValue)
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameter = new DataSourceParameter() { StatusId = statusIdToSet };

            // act
            var queryParameter = manager.GetParametersForQuery(parameter);

            // assert
            Assert.That(queryParameter, Is.Not.Null);
            Assert.That(queryParameter.Count, Is.EqualTo(1));
            Assert.That(queryParameter["status_id"], Is.EqualTo(expectedValue));
        }

        /// <summary>
        /// Test for the conversion of the parameter containing only an user id
        /// </summary>
        /// <param name="limitToSet">the id to set in the parameter</param>
        /// <param name="expectedValue">the value expected in the result</param>
        [TestCase(13, "13")]
        [TestCase(100, "100")]
        public void TestParameterConversionNoParameterButLimitSet(int limitToSet, string expectedValue)
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameter = new DataSourceParameter() { Limit = limitToSet };

            // act
            var queryParameter = manager.GetParametersForQuery(parameter);

            // assert
            Assert.That(queryParameter, Is.Not.Null);
            Assert.That(queryParameter.Count, Is.EqualTo(1));
            Assert.That(queryParameter["limit"], Is.EqualTo(expectedValue));
        }

        /// <summary>
        /// Test for the conversion of the parameter containing only an user id
        /// </summary>
        /// <param name="projectIdToSet">the id to set in the parameter</param>
        /// <param name="expectedValue">the value expected in the result</param>
        [TestCase(13, "13")]
        [TestCase(1, "1")]
        public void TestParameterConversionNoParameterButProjectIdSet(int projectIdToSet, string expectedValue)
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameter = new DataSourceParameter() { ProjectId = projectIdToSet };

            // act
            var queryParameter = manager.GetParametersForQuery(parameter);

            // assert
            Assert.That(queryParameter, Is.Not.Null);
            Assert.That(queryParameter.Count, Is.EqualTo(1));
            Assert.That(queryParameter["project_id"], Is.EqualTo(expectedValue));
        }

        /// <summary>
        /// Test for the conversion of the parameter containing only an user id
        /// </summary>
        [Test]
        public void TestParameterConversionNoParameterButSpentTimeSet()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameter = new DataSourceParameter() { SpentDateTimeTuple = Tuple.Create(new DateTime(2016, 1, 3), new DateTime(2017, 7, 26)) };

            // act
            var queryParameter = manager.GetParametersForQuery(parameter);

            // assert
            Assert.That(queryParameter, Is.Not.Null);
            Assert.That(queryParameter.Count, Is.EqualTo(1));
            Assert.That(queryParameter["spent_on"], Is.EqualTo("><2016-01-03|2017-07-26"));
        }

        /// <summary>
        /// Test for the conversion of the parameter containing only an user id
        /// </summary>
        [Test]
        public void TestParameterConversionNoParameterButUpdateStartTimeSet()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var parameter = new DataSourceParameter() { UpdateStartDateTime = new DateTime(2017, 2, 8) };

            // act
            var queryParameter = manager.GetParametersForQuery(parameter);

            // assert
            Assert.That(queryParameter, Is.Not.Null);
            Assert.That(queryParameter.Count, Is.EqualTo(1));
            Assert.That(queryParameter["updated_on"], Is.EqualTo(">=2017-02-08"));
        }

        /// <summary>
        /// Method to test the method to get the user
        /// </summary>
        [Test]
        public void TestUser()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            
            // act
            var user = manager.GetCurrentUser();

            // assert
            Assert.That(user, Is.Not.Null);
            Assert.That(user.Id, Is.Not.Null);
            Assert.That(user.Name, Contains.Substring("corpio"));
        }

        /// <summary>
        /// Integration test testing if creating and deleting a time entry works as expected
        /// </summary>
        [Test]
        public void CreateAndDeleteTimeEntry()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var activityInfo = manager.GetTotalActivityInfoList(new DataSourceParameter()).First(a => a.IsDefault);
            var issueInfo = manager.GetIssueInfoList(new DataSourceParameter() { Limit = 1 }).First();
            var projectInfo = new ProjectInfo() { Id = issueInfo.ProjectId, Name = issueInfo.ProjectShortName };

            var timeEntryInfo = new TimeEntryInfo()
                                    {
                                        Id = -5,
                                        Name = "CreateAndDeleteTest",
                                        StartDateTime = new DateTime(2016, 7, 8, 10, 0, 0),
                                        EndDateTime = new DateTime(2016, 7, 8, 12, 15, 0, 0),
                                        UpdateTime = new DateTime(2017, 2, 14),
                                        ActivityInfo = activityInfo,
                                        IssueInfo = issueInfo,
                                        ProjectInfo = projectInfo,
                                    };

            // act
            var created = manager.CreateObject(timeEntryInfo);
            var createdId = created.Id;
            
            // assert
            Assert.That(createdId, Is.Not.Null);

            // act
            manager.DeleteTimeEntry(createdId, new DataSourceParameter());
        }

        /// <summary>
        /// Test for the update of an issue
        /// </summary>
        [Test]
        public void TestUpdate()
        {
            // arrange
            var manager = new RedmineManagerInstance(Url, ApiKey);
            var activityInfo = manager.GetTotalActivityInfoList(new DataSourceParameter()).First(a => a.IsDefault);
            var issueInfo = manager.GetIssueInfoList(new DataSourceParameter() { Limit = 1 }).First();
            var projectInfo = new ProjectInfo() { Id = issueInfo.ProjectId, Name = issueInfo.ProjectShortName };

            var timeEntryInfo = new TimeEntryInfo()
            {
                Id = -5,
                Name = "CreateAndDeleteTest",
                StartDateTime = new DateTime(2016, 7, 8, 10, 0, 0),
                EndDateTime = new DateTime(2016, 7, 8, 12, 15, 0, 0),
                UpdateTime = new DateTime(2017, 2, 14),
                ActivityInfo = activityInfo,
                IssueInfo = issueInfo,
                ProjectInfo = projectInfo,
            };

            var created = manager.CreateObject(timeEntryInfo);
            var createdId = created.Id;

            // act
            var currentTime = DateTime.Now;
            timeEntryInfo.StartDateTime = new DateTime(2016, 7, 8, 8, 15, 0);
            manager.UpdateObject(timeEntryInfo);
            var updatedIssues =
                manager.GetTotalTimeEntryInfoList(new DataSourceParameter() { UpdateStartDateTime = currentTime })
                    .Where(i => i.Name == timeEntryInfo.Name);

            // clean up
            manager.DeleteTimeEntry(timeEntryInfo.Id, new DataSourceParameter());

            // assert
            Assert.That(updatedIssues, Is.Not.Null);
            Assert.That(updatedIssues.Count, Is.EqualTo(1));
            var issue = updatedIssues.First();
            Assert.That(issue.UpdateTime - DateTime.Now, Is.LessThanOrEqualTo(TimeSpan.FromMinutes(5)));
            Assert.That(issue.StartDateTime, Is.EqualTo(timeEntryInfo.StartDateTime));

            
        }

        #endregion
    }
}