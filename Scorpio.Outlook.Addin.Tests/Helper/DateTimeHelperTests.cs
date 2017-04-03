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

namespace Scorpio.Outlook.Addin.Tests.Helper
{
    using System;

    using NUnit.Framework;

    using Scorpio.Outlook.AddIn.Helper;

    /// <summary>
    /// Test class for testing the date time helper class
    /// </summary>
    [TestFixture]
    public class DateTimeHelperTests
    {
        /// <summary>
        /// Method to test the start of the month
        /// </summary>
        /// <param name="year">the year</param>
        /// <param name="month">the month</param>
        /// <param name="day">the day</param>
        [TestCase(2017, 2, 14)]
        [TestCase(2017, 2, 15)]
        [TestCase(2017, 2, 16)]
        [TestCase(2017, 2, 17)]
        [TestCase(2017, 2, 18)]
        [TestCase(2017, 2, 19)]
        [TestCase(2017, 2, 20)]
        [TestCase(2017, 2, 1)]
        [TestCase(2017, 2, 28)]
        [TestCase(2017, 1, 31)]
        [TestCase(2017, 1, 1)]

        public void StartOfMonthTest(int year, int month, int day)
        {
            // arrange
            var dateTime = new DateTime(year, month, day);

            // act
            var start = DateTimeHelper.StartOfMonth(dateTime);

            // assert
            Assert.That(start, Is.LessThanOrEqualTo(dateTime));
            Assert.That(dateTime - start, Is.LessThan(TimeSpan.FromDays(32)));
            Assert.That(dateTime.Date - start.Date, Is.LessThanOrEqualTo(TimeSpan.FromDays(31)));
            Assert.That(start.Date.Day, Is.EqualTo(1));
            Assert.That(start.Date.Month, Is.EqualTo(month));
            Assert.That(start.Date.Year, Is.EqualTo(year));
            Assert.That(start.Date, Is.EqualTo(start));
        }

        /// <summary>
        /// Method to test the end of the week
        /// </summary>
        /// <param name="year">the year</param>
        /// <param name="month">the month</param>
        /// <param name="day">the day</param>
        [TestCase(2017, 2, 14)]
        [TestCase(2017, 2, 15)]
        [TestCase(2017, 2, 16)]
        [TestCase(2017, 2, 17)]
        [TestCase(2017, 2, 18)]
        [TestCase(2017, 2, 19)]
        [TestCase(2017, 2, 20)]
        [TestCase(2017, 2, 1)]
        [TestCase(2017, 2, 28)]
        [TestCase(2017, 1, 31)]
        [TestCase(2017, 1, 1)]
        [TestCase(2017, 12, 31)]

        public void StartOfNextMonthTest(int year, int month, int day)
        {
            // arrange
            var dateTime = new DateTime(year, month, day);

            // act
            var nextMonthStart = DateTimeHelper.StartOfNextMonth(dateTime);

            // assert
            Assert.That(nextMonthStart, Is.GreaterThan(dateTime));
            Assert.That(dateTime - nextMonthStart, Is.LessThan(TimeSpan.FromDays(32)));
            Assert.That(dateTime.Date - nextMonthStart.Date, Is.LessThanOrEqualTo(TimeSpan.FromDays(31)));
            Assert.That(nextMonthStart.Date.Day, Is.EqualTo(1));
            Assert.That(nextMonthStart.Date.Year - year, Is.LessThanOrEqualTo(1));
            Assert.That(nextMonthStart.Date, Is.EqualTo(nextMonthStart));
            Assert.That(Math.Abs((month % 12) - (nextMonthStart.Month % 12)), Is.EqualTo(1));
        }

        /// <summary>
        /// Method to test the start of the week
        /// </summary>
        /// <param name="year">the year</param>
        /// <param name="month">the month</param>
        /// <param name="day">the day</param>
        [TestCase(2017, 2, 14)]
        [TestCase(2017, 2, 15)]
        [TestCase(2017, 2, 16)]
        [TestCase(2017, 2, 17)]
        [TestCase(2017, 2, 18)]
        [TestCase(2017, 2, 19)]
        [TestCase(2017, 2, 20)]

        public void StartOfWeekTest(int year, int month, int day)
        {
            // arrange
            var dateTime = new DateTime(year, month, day);

            // act
            var start = DateTimeHelper.StartOfWeek(dateTime);

            // assert
            Assert.That(start, Is.LessThanOrEqualTo(dateTime));
            Assert.That(dateTime - start, Is.LessThan(TimeSpan.FromDays(7)));
            Assert.That(start.DayOfWeek, Is.EqualTo(DayOfWeek.Monday));
        }

        /// <summary>
        /// Method to test the end of the week
        /// </summary>
        /// <param name="year">the year</param>
        /// <param name="month">the month</param>
        /// <param name="day">the day</param>
        [TestCase(2017, 2, 14)]
        [TestCase(2017, 2, 15)]
        [TestCase(2017, 2, 16)]
        [TestCase(2017, 2, 17)]
        [TestCase(2017, 2, 18)]
        [TestCase(2017, 2, 19)]
        [TestCase(2017, 2, 20)]

        public void StartOfNextWeekTest(int year, int month, int day)
        {
            // arrange
            var dateTime = new DateTime(year, month, day);

            // act
            var nextStartOfWeek = DateTimeHelper.StartOfNextWeek(dateTime);

            // assert
            Assert.That(nextStartOfWeek, Is.GreaterThanOrEqualTo(dateTime));
            Assert.That(nextStartOfWeek - dateTime, Is.LessThanOrEqualTo(TimeSpan.FromDays(7)));
            Assert.That(nextStartOfWeek.DayOfWeek, Is.EqualTo(DayOfWeek.Monday));
            Assert.That(nextStartOfWeek.Date, Is.EqualTo(nextStartOfWeek));
        }

        /// <summary>
        /// Method to test the is workday function
        /// </summary>
        /// <param name="year">the year</param>
        /// <param name="month">the month</param>
        /// <param name="day">the day</param>
        /// <param name="isWorkdayExpected">if the date is expected to be a workday</param>
        [TestCase(2017, 2, 14, true)]
        [TestCase(2017, 2, 15, true)]
        [TestCase(2017, 2, 16, true)]
        [TestCase(2017, 2, 17, true)]
        [TestCase(2017, 2, 18, false)]
        [TestCase(2017, 2, 19, false)]
        [TestCase(2017, 2, 20, true)]
        [TestCase(2017, 2, 21, true)]
        public void IsWorkdayTest(int year, int month, int day, bool isWorkdayExpected)
        {
            // arrange
            var dateTime = new DateTime(year, month, day);

            // act
            var isWorkday = DateTimeHelper.IsWorkDay(dateTime);

            // assert
            Assert.That(isWorkday, Is.EqualTo(isWorkdayExpected));
        }

        /// <summary>
        /// Test for the start of the day
        /// </summary>
        /// <param name="year">the year</param>
        /// <param name="month">the month</param>
        /// <param name="day">the day</param>
        /// <param name="hour">the hour</param>
        /// <param name="minute">the minute</param>
        [TestCase(2017, 2, 21, 12, 7)]
        [TestCase(2017, 2, 21, 0, 0)]
        [TestCase(2017, 2, 21, 17, 1)]
        [TestCase(2017, 2, 21, 3, 5)]
        [TestCase(2017, 2, 21, 23, 59)]
        public void StartOfDayTest(int year, int month, int day, int hour, int minute)
        {
            // arrange
            var dateTime = new DateTime(year, month, day, hour, minute, 0);

            // act
            var start = DateTimeHelper.StartOfDay(dateTime);

            // assert
            Assert.That(start, Is.LessThanOrEqualTo(dateTime));
            Assert.That(dateTime - start, Is.LessThan(TimeSpan.FromDays(1)));
            Assert.That(start.Date, Is.EqualTo(dateTime.Date));
        }

        /// <summary>
        /// Test for the start of the day
        /// </summary>
        /// <param name="year">the year</param>
        /// <param name="month">the month</param>
        /// <param name="day">the day</param>
        /// <param name="hour">the hour</param>
        /// <param name="minute">the minute</param>
        [TestCase(2017, 2, 21, 12, 7)]
        [TestCase(2017, 2, 21, 0, 0)]
        [TestCase(2017, 2, 21, 17, 1)]
        [TestCase(2017, 2, 21, 3, 5)]
        [TestCase(2017, 2, 21, 23, 59)]
        public void StartOfNextDayTest(int year, int month, int day, int hour, int minute)
        {
            // arrange
            var dateTime = new DateTime(year, month, day, hour, minute, 0);

            // act
            var startOfNextDay = DateTimeHelper.StartOfNextDay(dateTime);

            // assert
            Assert.That(startOfNextDay, Is.GreaterThanOrEqualTo(dateTime));
            Assert.That(startOfNextDay - dateTime, Is.LessThanOrEqualTo(TimeSpan.FromDays(1)));
            Assert.That(startOfNextDay.Date, Is.EqualTo(startOfNextDay));
        }

        /// <summary>
        /// Test to get the minimum date time
        /// </summary>
        [Test]
        public void TestMinimumDateTime()
        {
            // arrange
            var dateTime1 = new DateTime(2016, 7, 8);
            var dateTime2 = new DateTime(2016, 7, 8, 12, 11, 10);
            var dateTime3 = new DateTime(2016, 8, 8, 12, 11, 10);

            // act
            var minimum = DateTimeHelper.MinimumDateTime(dateTime1, dateTime2, dateTime3);

            // assert
            Assert.That(minimum, Is.EqualTo(dateTime1));
        }

        /// <summary>
        /// Test to get the maximum date time
        /// </summary>
        [Test]
        public void TestMaximumDateTime()
        {
            // arrange
            var dateTime1 = new DateTime(2016, 7, 8);
            var dateTime2 = new DateTime(2016, 7, 8, 12, 11, 10);
            var dateTime3 = new DateTime(2016, 8, 8, 12, 11, 10);

            // act
            var minimum = DateTimeHelper.MaximumDateTime(dateTime1, dateTime2, dateTime3);

            // assert
            Assert.That(minimum, Is.EqualTo(dateTime3));
        }
    }
}