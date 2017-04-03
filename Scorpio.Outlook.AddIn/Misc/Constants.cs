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

namespace Scorpio.Outlook.AddIn.Misc
{
    /// <summary>
    /// Class contains the constants of the plugin.
    /// </summary>
    public class Constants
    {
        #region Constants

        /// <summary>
        /// Timer interval for synchronization.
        /// </summary>
        public const int TimerInterval = 2000;

        /// <summary>
        /// Epsilon constant which is used in equality checks of float values.
        /// </summary>
        public const double Epsilon = 0.0001;

        /// <summary>
        /// The name of the RedmineTimeEntryId field on outlook appointments
        /// </summary>
        public const string FieldRedmineTimeEntryId = "RedmineTimeEntryID";

        /// <summary>
        /// The name of the RedmineProjectId field on outlook appointments
        /// </summary>
        public const string FieldRedmineProjectId = "RedmineProjectID";

        /// <summary>
        /// The name of the RedmineIssueId field on outlook appointments
        /// </summary>
        public const string FieldRedmineIssueId = "RedmineIssueID";

        /// <summary>
        /// The name of the RedmineActivityId field on outlook appointments
        /// </summary>
        public const string FieldRedmineActivityId = "RedmineActivityID";

        /// <summary>
        /// The name of the field on outlook appointments that determines from which appointment the appointment was copied
        /// </summary>
        public const string FieldEntryIdCopy = "CopyOfEntryID";

        /// <summary>
        /// The name of the field on outlook appointments that determines which AppointmentState the appointment has
        /// </summary>
        public const string FieldAppointmentState = "AppointmentState";

        /// <summary>
        /// The name of the field on outlook appointments that determines which AppointmentState the appointment had 
        /// before its previous state. Needed when an appointment state is set to sync-Error, and we want to reapply 
        /// the correct synchronization again later.
        /// </summary>
        public const string FieldAppointmentPreviousState = "AppointmentPreviousState";

        /// <summary>
        /// The name of the RedmineLastUpdate field on outlook appointments
        /// </summary>
        public const string FieldLastUpdate = "RedmineLastUpdate";

        /// <summary>
        /// The name of the FieldImportedFromRedmine field on outlook appointments. If this field is set, it means that the appointment has
        /// just been imported from redmine.
        /// </summary>
        public const string FieldImportedFromRedmine = "ImportedFromRedmine";

        /// <summary>
        /// The amount of minutes in an hour
        /// </summary>
        public const int MinutesInHour = 60;

        #endregion
    }
}