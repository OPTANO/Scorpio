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

namespace Scorpio.Outlook.AddIn.Misc
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;

    using Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Type safe "enum" that declares possible appointment states that can be applied to AppointmentItem instances. These states are realized as Outlook Categories.
    /// </summary>
    public class AppointmentState
    {
        #region Static Fields

        /// <summary>
        /// State that marks an appointment as deleted. This means that it is still in the outlook calendar, but will be removed from redmine and the calendar in the next synchronization attempt.
        /// </summary>
        public static readonly AppointmentState Deleted = new AppointmentState(4, "ORC Deleted", OlCategoryColor.olCategoryColorDarkPurple);

        /// <summary>
        /// State that marks an appointment as modified. This means any of its properties have been changed and are not synchronized with redmine yet.
        /// </summary>
        public static readonly AppointmentState Modified = new AppointmentState(2, "ORC Modified", OlCategoryColor.olCategoryColorDarkOrange);

        /// <summary>
        /// State that marks an appointment as having a synchronization error. This means something went wrong in the last synchronization attempt.
        /// </summary>
        public static readonly AppointmentState SyncError = new AppointmentState(3, "ORC Sync Error", OlCategoryColor.olCategoryColorDarkRed);

        /// <summary>
        /// State that marks an appointment as synchronized. This means that the appointment has no changes made to it since the last successful synchronization.
        /// </summary>
        public static readonly AppointmentState Synchronized = new AppointmentState(1, "ORC Synced", OlCategoryColor.olCategoryColorDarkGreen);

        /// <summary>
        /// State that marks an overtime appointment as synchronized. This means that the appointment has no changes made to it since the last successful synchronization.
        /// </summary>
        public static readonly AppointmentState SynchronizedOvertime = new AppointmentState(
                                                                           5,
                                                                           "ORC Synced Overtime",
                                                                           OlCategoryColor.olCategoryColorBlue);

        /// <summary>
        /// All AppointmentStates that are defined.
        /// </summary>
        public static readonly IEnumerable<AppointmentState> AllStates =
            new ReadOnlyCollection<AppointmentState>(new List<AppointmentState> { Synchronized, SynchronizedOvertime, Modified, SyncError, Deleted });

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="AppointmentState"/> class.
        /// </summary>
        /// <param name="value">The value</param>
        /// <param name="name">The name</param>
        /// <param name="color">The color</param>
        private AppointmentState(int value, string name, OlCategoryColor color)
        {
            this.Name = name;
            this.Value = value;
            this.Color = color;
        }

        #endregion
        
        #region Public properties

        /// <summary>
        /// Gets the name of the state
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the integer value of the state
        /// </summary>
        public int Value { get; private set; }

        /// <summary>
        /// Gets the outlook category color of the state
        /// </summary>
        public OlCategoryColor Color { get; private set; }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Checks if a specified string matches any of the appointmentstates names.
        /// </summary>
        /// <param name="name">The name to check</param>
        /// <returns>True if there is an appointment state that has the same name as the provided parameter. False otherwise.</returns>
        public static bool IsValidAppointmentStateName(string name)
        {
            return AllStates.Any(state => state.Name == name);
        }

        /// <summary>
        /// ToString implementation
        /// </summary>
        /// <returns>The name of the appointmentstate</returns>
        public override string ToString()
        {
            return this.Name;
        }

        #endregion

    }
}