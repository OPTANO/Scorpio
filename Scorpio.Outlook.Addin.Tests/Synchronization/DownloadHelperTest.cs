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

namespace Scorpio.Outlook.Addin.Tests.Synchronization
{
    using System.Collections.Generic;
    using System.Linq;

    using NUnit.Framework;

    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Synchronization.ExternalDataSource;
    using Scorpio.Outlook.AddIn.Synchronization.Helper;

    /// <summary>
    /// Test class containing tests for the synchronizer class
    /// </summary>
    [TestFixture]
    public class DownloadHelperTest
    {
        /// <summary>
        /// Method to test the download of issues
        /// </summary>
        [Test]
        public void TestIssueDownloader()
        {
            // initialize parameter needed
            var newProjects = new List<ProjectInfo>();
            var currentIssues = new Dictionary<int, IssueInfo>();

            // get external data source
            ExternalDataSourceFactory.UseTestManager = true;
            var source = ExternalDataSourceFactory.GetRedmineMangerInstance("", "", 50);
            
            // download issues
            var issues = DownloadHelper.DownloadIssues(newProjects, currentIssues, source);

            // assert
            Assert.That(issues.Count, Is.GreaterThan(0));
        }

        /// <summary>
        /// Method to test the download of issues
        /// </summary>
        [Test]
        public void TestNewIssuesOverExistingIssues()
        {
            // initialize parameter needed
            var newProjects = new List<ProjectInfo>();
            var currentIssues = new Dictionary<int, IssueInfo>();
            var changedIssue = new IssueInfo() { Id = 0, ProjectId = 42 };
            currentIssues.Add(0, changedIssue);
            
            // get external data source
            ExternalDataSourceFactory.UseTestManager = true;
            var source = ExternalDataSourceFactory.GetRedmineMangerInstance("", "", 50);
            var issuesChangedProject = DownloadHelper.DownloadIssues(newProjects, currentIssues, source);
            
            // assert
            Assert.That(issuesChangedProject.Count, Is.GreaterThanOrEqualTo(1));
            Assert.That(changedIssue.Id.HasValue, Is.True);
            var issueId = changedIssue.Id.Value;
            Assert.That(issuesChangedProject.Keys, Contains.Item(issueId));
            var issue = issuesChangedProject[issueId];
            Assert.That(issue.ProjectId, Is.EqualTo(0));
        }
    }
}