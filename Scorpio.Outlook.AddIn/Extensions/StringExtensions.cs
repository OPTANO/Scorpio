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

namespace Scorpio.Outlook.AddIn.Extensions
{
    /// <summary>
    /// Extension methods for strings
    /// </summary>
    public static class StringExtensions
    {
        #region Constants

        /// <summary>
        /// The minimum length for an issue to be checked
        /// </summary>
        private const int MinLength = 5;

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Gets teh string to use for search of unknown issues
        /// </summary>
        /// <param name="text">the text to check</param>
        /// <returns>the text to use for search or null if none should be used</returns>
        public static string GetStringToUseForUnknownIssueSearch(this string text)
        {
            string stringToReturn = null;

            if (text != null)
            {
                // if the string start with #, always check for new issue
                var startsWithHashtag = text.StartsWith("#");
                if (startsWithHashtag)
                {
                    stringToReturn = text.Substring(1);
                }
                else
                {
                    // if the issue does not start with #, check if it contains at least 5 digits
                    var length = text.Length;
                    if (length >= MinLength)
                    {
                        stringToReturn = startsWithHashtag ? text.Substring(1) : text;
                    }
                }
            }
            return stringToReturn;
        }

        #endregion
    }
}