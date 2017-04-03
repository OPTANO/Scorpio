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

namespace Scorpio.Outlook.AddIn.Helper
{
    using System;

    using Microsoft.Office.Interop.Outlook;

    using Scorpio.Outlook.AddIn.Misc;

    /// <summary>
    /// Class that provides helper methods for dealing with Outlook datastructures.
    /// </summary>
    public class OutlookHelper
    {
        #region Public Methods and Operators

        /// <summary>
        /// Creates or gets a user property field
        /// </summary>
        /// <param name="folder">The target folder</param>
        /// <param name="name">The name</param>
        /// <param name="type">The property value type</param>
        /// <returns>The user property field</returns>
        public static UserDefinedProperty CreateOrGetProperty(MAPIFolder folder, string name, OlUserPropertyType type)
        {
            foreach (var prop in folder.UserDefinedProperties)
            {
                if (((UserDefinedProperty)prop).Name == name)
                {
                    return (UserDefinedProperty)prop;
                }
            }
            return folder.UserDefinedProperties.Add(name, type, true);
        }

        /// <summary>
        /// Creates or gets a folder
        /// </summary>
        /// <param name="parent">The parent folder</param>
        /// <param name="name">The name of the new folder</param>
        /// <param name="newType">The type of the new folder</param>
        /// <returns>The new folder</returns>
        public static MAPIFolder CreateOrGetFolder(MAPIFolder parent, string name, OlDefaultFolders newType)
        {
            foreach (Folder personalFolder in parent.Folders)
            {
                if (personalFolder.Name == name)
                {
                    return personalFolder;
                }
            }
            return parent.Folders.Add(name, newType);
        }

        /// <summary>
        /// Creates or gets a user property field
        /// </summary>
        /// <param name="item">The target item</param>
        /// <param name="name">The name</param>
        /// <param name="type">The property value type</param>
        /// <returns>The user property field</returns>
        public static UserProperty CreateOrGetProperty(AppointmentItem item, string name, OlUserPropertyType type)
        {
            foreach (var prop in item.UserProperties)
            {
                if (((UserProperty)prop).Name == name)
                {
                    return (UserProperty)prop;
                }
            }
            return item.UserProperties.Add(name, type, true);
        }

        /// <summary>
        /// Checks whether a categorie exists
        /// </summary>
        /// <param name="categoryName">The name of the category</param>
        /// <returns>True if the category exists</returns>
        public static bool CategoryExists(string categoryName)
        {
            try
            {
                var category = Globals.ThisAddIn.Application.Session.Categories[categoryName];
                return category != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Method that creates all the necessary appointment categories that are needed for scorpio.
        /// </summary>
        public static void CreateScorpioCategories()
        {
            // create the categories if needed
            var categories = Globals.ThisAddIn.Application.Session.Categories;
            foreach (var category in AppointmentState.AllStates)
            {
                if (!CategoryExists(category.Name))
                {
                    categories.Add(category.Name, category.Color);
                }
            }
        }

        /// <summary>
        /// Method that registers user defined properties to the redmine calendar folder.
        /// </summary>
        /// <param name="redmineTimeEntriesFolder">The folder which contains the redmine time entries appointments.</param>
        public static void CreateScorpioUserDefinedProperties(MAPIFolder redmineTimeEntriesFolder)
        {
            redmineTimeEntriesFolder.UserDefinedProperties.Add(
                Constants.FieldAppointmentPreviousState, 
                OlUserPropertyType.olInteger, 
                Type.Missing, 
                Type.Missing);
            redmineTimeEntriesFolder.UserDefinedProperties.Add(
                Constants.FieldAppointmentState, 
                OlUserPropertyType.olInteger, 
                Type.Missing, 
                Type.Missing);
            redmineTimeEntriesFolder.UserDefinedProperties.Add(Constants.FieldEntryIdCopy, OlUserPropertyType.olText, Type.Missing, Type.Missing);
            redmineTimeEntriesFolder.UserDefinedProperties.Add(Constants.FieldLastUpdate, OlUserPropertyType.olDateTime, Type.Missing, Type.Missing);
            redmineTimeEntriesFolder.UserDefinedProperties.Add(
                Constants.FieldRedmineActivityId, 
                OlUserPropertyType.olInteger, 
                Type.Missing, 
                Type.Missing);
            redmineTimeEntriesFolder.UserDefinedProperties.Add(
                Constants.FieldRedmineIssueId, 
                OlUserPropertyType.olInteger, 
                Type.Missing, 
                Type.Missing);
            redmineTimeEntriesFolder.UserDefinedProperties.Add(
                Constants.FieldRedmineProjectId, 
                OlUserPropertyType.olInteger, 
                Type.Missing, 
                Type.Missing);
            redmineTimeEntriesFolder.UserDefinedProperties.Add(
                Constants.FieldRedmineTimeEntryId, 
                OlUserPropertyType.olInteger, 
                Type.Missing, 
                Type.Missing);
        }

        #endregion
    }
}