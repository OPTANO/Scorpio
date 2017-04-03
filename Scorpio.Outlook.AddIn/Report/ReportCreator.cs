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

namespace Scorpio.Outlook.AddIn.Report
{
    using System;
    using System.Collections.Generic;

    using Microsoft.Office.Interop.Outlook;

    using Scorpio.Outlook.AddIn.Helper;
    using Scorpio.Outlook.AddIn.Misc;

    /// <summary>
    /// Class that handles the generation and preparation of the monthly report.
    /// </summary>
    public class ReportCreator
    {
        #region Static Fields

        /// <summary>
        /// The name of the note that contains contracts.
        /// </summary>
        public static readonly string ContractsName = "Verträge";

        /// <summary>
        /// The name of the note that contains month sheets.
        /// </summary>
        public static readonly string MonthSheetsName = "Stundenzettel";

        #endregion

        #region Public properties

        /// <summary>
        /// Gets the contracts folder
        /// </summary>
        public MAPIFolder ContractsFolder { get; private set; }

        /// <summary>
        /// Gets the month sheets folder
        /// </summary>
        public MAPIFolder MonthSheetsFolder { get; private set; }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Method that is called for initialization. It checks if all necessary items for the monthly report 
        /// are present in the current outlook account. If those items are not present, they are created.
        /// </summary>
        /// <param name="currentExplorer">The current explorer</param>
        public void CheckRequirements(Explorer currentExplorer)
        {
            // create the contracts and month sheet folders
            var primaryNotes = currentExplorer.Session.GetDefaultFolder(OlDefaultFolders.olFolderNotes);
            this.ContractsFolder = OutlookHelper.CreateOrGetFolder(primaryNotes, ContractsName, OlDefaultFolders.olFolderNotes);
            this.MonthSheetsFolder = OutlookHelper.CreateOrGetFolder(primaryNotes, MonthSheetsName, OlDefaultFolders.olFolderNotes);

            // Create Field in Contracts Folder
            OutlookHelper.CreateOrGetProperty(this.ContractsFolder, "Startdatum", OlUserPropertyType.olDateTime);
            OutlookHelper.CreateOrGetProperty(this.ContractsFolder, "Enddatum", OlUserPropertyType.olDateTime);
            OutlookHelper.CreateOrGetProperty(this.ContractsFolder, "Wöchenliche Arbeitszeit", OlUserPropertyType.olNumber);
            OutlookHelper.CreateOrGetProperty(this.ContractsFolder, "Urlaubstage", OlUserPropertyType.olNumber);
            OutlookHelper.CreateOrGetProperty(this.ContractsFolder, "Startsaldo", OlUserPropertyType.olNumber);

            // Create Field in MonthSheets Folder
            OutlookHelper.CreateOrGetProperty(this.MonthSheetsFolder, "Monat", OlUserPropertyType.olText);
            OutlookHelper.CreateOrGetProperty(this.MonthSheetsFolder, "Soll", OlUserPropertyType.olNumber);
            OutlookHelper.CreateOrGetProperty(this.MonthSheetsFolder, "Ist", OlUserPropertyType.olNumber);
            OutlookHelper.CreateOrGetProperty(this.MonthSheetsFolder, "Saldo", OlUserPropertyType.olNumber);
            OutlookHelper.CreateOrGetProperty(this.MonthSheetsFolder, "Urlaub genommen", OlUserPropertyType.olNumber);
            OutlookHelper.CreateOrGetProperty(this.MonthSheetsFolder, "Resturlaub", OlUserPropertyType.olNumber);
        }

        /// <summary>
        /// Method that creates the monthly report.
        /// </summary>
        /// <param name="synchronizer">
        /// The redmine synchronizer
        /// </param>
        /// <param name="application">
        /// The application.
        /// </param>
        public async void CalculateMonthlyReport(Synchronizer synchronizer, Application application)
        {
            // TODO DS: reenable writing of report!

            ////var datetime = Globals.ThisAddIn.CalendarState.GetCurrentMonth();
            ////if (
            ////    MessageBox.Show(
            ////        string.Format("Stundenzettel für {0:D2}/{1:D4} berechnen?", datetime.Month, datetime.Year),
            ////        "ORCONOMY Tool",
            ////        MessageBoxButtons.YesNo) == DialogResult.No)
            ////{
            ////    return;
            ////}

            ////var currentMonth = await synchronizer.PrepareCurrentMonthForEvaluation();

            ////// Calculate Workdays for Month
            ////var firstDayofMonth = new DateTime(datetime.Year, datetime.Month, 1);
            ////var lastDayofMonth = firstDayofMonth.AddMonths(1).Subtract(new TimeSpan(1, 0, 0, 0, 0));
            ////var workdays = this.CountWeekdays(firstDayofMonth, lastDayofMonth);

            ////// Calculate Daily Hours
            ////// TODO: Auslesen/Berechnen
            ////var dailyHours = this.CalculateDailyHours(firstDayofMonth, lastDayofMonth, workdays);

            ////// Load Appointments
            ////var appointmentsItems = currentMonth.Item2; // Synchronizer.GetAppointmentsInRange(firstDayofMonth, lastDayofMonth);

            ////// Check for Holidays
            ////// TODO

            ////// Calculate Work Balance
            ////var loggedTime = 0.0;

            ////foreach (var appointment in appointmentsItems)
            ////{
            ////    loggedTime += (appointment.End - appointment.Start).TotalHours;
            ////}
            ////// TODO Saldo auslesen

            ////// Saldo des vorherigen Monats holen berechnen
            ////var saldo = this.CalculateSaldoBeforeMonth(datetime, firstDayofMonth, lastDayofMonth);

            ////// Write Report
            ////var targetTime = dailyHours * workdays;
            ////saldo = saldo + loggedTime - targetTime;

            ////NoteItem monthsheet =
            ////    this.GetMonthSheet(
            ////        string.Format(
            ////            "{0} {1}",
            ////            System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(datetime.Month),
            ////            datetime.Year));

            ////foreach (ItemProperty itemProperty in monthsheet.ItemProperties)
            ////{
            ////    if (itemProperty.Name == "Soll")
            ////    {
            ////        itemProperty.Value = targetTime;
            ////    }

            ////    if (itemProperty.Name == "Ist")
            ////    {
            ////        itemProperty.Value = loggedTime;
            ////    }

            ////    if (itemProperty.Name == "Saldo")
            ////    {
            ////        itemProperty.Value = saldo;
            ////    }
            ////}

            ////monthsheet.Save();

            ////// Write Mail
            ////var mailItem = (MailItem)application.CreateItem(OlItemType.olMailItem);
            ////mailItem.Subject = "Stundenreport "
            ////                   + System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(
            ////                       datetime.Month) + " " + datetime.Year;
            ////mailItem.To = "";
            ////mailItem.HTMLBody = string.Format("<h2>Stundenreport {0} {1} </h2> Mitarbeiter: {2} <br/> Soll (Stunden): {3} Stunden <br/> Ist (Stunden): {4} Stunden <br/> Saldo: <b>{5}</b>", System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(datetime.Month), datetime.Year, synchronizer.CurrentUserName, targetTime, loggedTime, saldo);
            ////mailItem.Display(true);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Calculates the salde before a specified month.
        /// </summary>
        /// <param name="datetime">The month for which to calculate the saldo up to this month.</param>
        /// <param name="firstDayofMonth">The first day of month.</param>
        /// <param name="lastDayofMonth">The last day of month.</param>
        /// <returns>The saldo.</returns>
        private double CalculateSaldoBeforeMonth(DateTime datetime, DateTime firstDayofMonth, DateTime lastDayofMonth)
        {
            var saldo = 0.0;
            NoteItem lastmonthsheet = null;

            if (datetime.Month != 1)
            {
                var monthsheets = this.MonthSheetsFolder.Items;

                foreach (NoteItem lastmonthsheet1 in monthsheets)
                {
                    foreach (ItemProperty itemProperty in lastmonthsheet1.ItemProperties)
                    {
                        if (itemProperty.Name == "Monat")
                        {
                            if (itemProperty.Value
                                == string.Format(
                                    "{0} {1}", 
                                    System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(datetime.Month - 1), 
                                    datetime.Year))
                            {
                                lastmonthsheet = lastmonthsheet1;
                            }
                        }
                    }
                }
            }
            else
            {
                var monthsheets = this.MonthSheetsFolder.Items;

                foreach (NoteItem lastmonthsheet1 in monthsheets)
                {
                    foreach (ItemProperty itemProperty in lastmonthsheet1.ItemProperties)
                    {
                        if (itemProperty.Name == "Monat")
                        {
                            if (itemProperty.Value
                                == string.Format(
                                    "{0} {1}", 
                                    System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(12), 
                                    datetime.Year - 1))
                            {
                                lastmonthsheet = lastmonthsheet1;
                            }
                        }
                    }
                }
            }

            if (lastmonthsheet != null)
            {
                foreach (ItemProperty itemProperty in lastmonthsheet.ItemProperties)
                {
                    if (itemProperty.Name == "Saldo")
                    {
                        saldo = itemProperty.Value;
                    }
                }
            }
            else
            {
                var contracts = this.ContractsFolder.Items;

                foreach (NoteItem contract in contracts)
                {
                    // var properties = contract.ItemProperties;
                    var startdate = DateTime.Now;
                    var enddate = DateTime.Now;
                    var startsaldo = 0.0;

                    foreach (ItemProperty property in contract.ItemProperties)
                    {
                        if (property.Name == "Startdatum")
                        {
                            startdate = property.Value;
                        }
                        if (property.Name == "Enddatum")
                        {
                            enddate = property.Value;
                        }
                        if (property.Name == "Startsaldo")
                        {
                            startsaldo = property.Value;
                        }
                    }

                    if (startdate < lastDayofMonth && enddate > firstDayofMonth)
                    {
                        saldo += startsaldo;
                    }
                }
            }
            return saldo;
        }

        /// <summary>
        /// Gets the monthsheet for a specified month.
        /// </summary>
        /// <param name="month">The month for which to get the sheet.</param>
        /// <returns>The <see cref="NoteItem"/> which represents the month sheet of the requested month.</returns>
        private NoteItem GetMonthSheet(string month)
        {
            var monthsheets = this.MonthSheetsFolder.Items;

            foreach (NoteItem monthsheet in monthsheets)
            {
                foreach (ItemProperty itemProperty in monthsheet.ItemProperties)
                {
                    if (itemProperty.Name == "Monat")
                    {
                        if (itemProperty.Value == month)
                        {
                            return monthsheet;
                        }
                    }
                }
            }

            NoteItem newmonthsheet = this.MonthSheetsFolder.Items.Add(OlItemType.olNoteItem);
            newmonthsheet.ItemProperties.Add("Monat", OlUserPropertyType.olText);
            newmonthsheet.ItemProperties.Add("Soll", OlUserPropertyType.olNumber);
            newmonthsheet.ItemProperties.Add("Ist", OlUserPropertyType.olNumber);
            newmonthsheet.ItemProperties.Add("Saldo", OlUserPropertyType.olNumber);
            newmonthsheet.ItemProperties.Add("Urlaub genommen", OlUserPropertyType.olNumber);
            newmonthsheet.ItemProperties.Add("Resturlaub", OlUserPropertyType.olNumber);

            foreach (ItemProperty itemProperty in newmonthsheet.ItemProperties)
            {
                if (itemProperty.Name == "Monat")
                {
                    itemProperty.Value = month;
                }
            }

            newmonthsheet.Save();

            return newmonthsheet;
        }

        /// <summary>
        /// Calculates the Daily Hours
        /// </summary>
        /// <param name="firstDayofMonth">The first Day of Month</param>
        /// <param name="lastDayofMonth">The last Day of Month</param>
        /// <param name="workdays">Workdays in Month</param>
        /// <returns>The Daily Hours</returns>
        private double CalculateDailyHours(DateTime firstDayofMonth, DateTime lastDayofMonth, int workdays)
        {
            var dailyHours = 0.0;
            var contracts = this.ContractsFolder.Items;
            var listOfValidContracts = new List<NoteItem>();

            foreach (NoteItem contract in contracts)
            {
                // var properties = contract.ItemProperties;
                DateTime startdate = DateTime.Now;
                DateTime enddate = DateTime.Now;

                foreach (ItemProperty property in contract.ItemProperties)
                {
                    if (property.Name == "Startdatum")
                    {
                        startdate = property.Value;
                    }
                    if (property.Name == "Enddatum")
                    {
                        enddate = property.Value;
                    }

                    // if (property.Name == "Wöchenliche Arbeitszeit")
                    // {
                    // hoursPerDay = property.Value/5;
                    // }
                }

                if (startdate < lastDayofMonth && enddate > firstDayofMonth)
                {
                    listOfValidContracts.Add(contract);
                }
            }

            if (listOfValidContracts.Count == 1)
            {
                foreach (ItemProperty property in listOfValidContracts[0].ItemProperties)
                {
                    if (property.Name == "Wöchenliche Arbeitszeit")
                    {
                        dailyHours = property.Value / 5;
                    }
                }
            }
            else
            {
                var hoursToWork = 0.0;

                foreach (var contract in listOfValidContracts)
                {
                    DateTime startdate = DateTime.Now;
                    DateTime enddate = DateTime.Now;
                    double hoursPerDay = 0;

                    foreach (ItemProperty property in contract.ItemProperties)
                    {
                        if (property.Name == "Startdatum")
                        {
                            startdate = (DateTime)property.Value.Date;
                        }
                        if (property.Name == "Enddatum")
                        {
                            enddate = (DateTime)property.Value.Date;
                        }
                        if (property.Name == "Wöchenliche Arbeitszeit")
                        {
                            hoursPerDay = property.Value / 5;
                        }
                    }

                    if (startdate > firstDayofMonth && enddate < lastDayofMonth)
                    {
                        hoursToWork += this.CountWeekdays(startdate, enddate) * hoursPerDay;
                    }
                    else
                    {
                        if (startdate > firstDayofMonth)
                        {
                            hoursToWork += this.CountWeekdays(startdate, lastDayofMonth) * hoursPerDay;
                        }

                        if (enddate < lastDayofMonth)
                        {
                            hoursToWork += this.CountWeekdays(firstDayofMonth, enddate) * hoursPerDay;
                        }
                    }
                }

                dailyHours = hoursToWork / workdays;
            }

            return dailyHours;
        }

        /// <summary>
        /// Counts the amount of weekdays between two dates.
        /// </summary>
        /// <param name="startTime">The start date</param>
        /// <param name="endTime">The end date</param>
        /// <returns>The amount of weekdays between start and end time.</returns>
        private int CountWeekdays(DateTime startTime, DateTime endTime)
        {
            TimeSpan timeSpan = endTime - startTime;
            DateTime dateTime;
            int weekdays = 0;
            for (int i = 0; i <= timeSpan.Days; i++)
            {
                dateTime = startTime.AddDays(i);
                if (this.IsWeekDay(dateTime))
                {
                    weekdays++;
                }
            }
            return weekdays;
        }

        /// <summary>
        /// Checks if a date is on Mon-Fri.
        /// </summary>
        /// <param name="dateTime">The date to check.</param>
        /// <returns><code>true</code> if the date is on a weekday, <code>false</code> if the date is on a weekend.</returns>
        private bool IsWeekDay(DateTime dateTime)
        {
            return (dateTime.DayOfWeek != DayOfWeek.Saturday) && (dateTime.DayOfWeek != DayOfWeek.Sunday);
        }

        #endregion
    }
}