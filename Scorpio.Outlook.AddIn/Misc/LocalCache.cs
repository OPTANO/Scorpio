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

namespace Scorpio.Outlook.AddIn.Cache
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.IO.IsolatedStorage;
    using System.Linq;
    using System.Runtime.Serialization.Formatters.Binary;

    using log4net;

    using Scorpio.Outlook.AddIn.LocalObjects;
    using Scorpio.Outlook.AddIn.Properties;

    /// <summary>
    /// Class that manages access to the isolated storage for the application.
    /// </summary>
    public class LocalCache
    {
        #region Constants

        /// <summary>
        /// The key by which the data for known issues can be accessed.
        /// </summary>
        public const string KnownIssues = "knownissues";

        /// <summary>
        /// The key by which the data for known projects can be accessed.
        /// </summary>
        public const string KnownProjects = "knownprojects";

        #endregion

        #region Static Fields

        /// <summary>
        /// The logger.
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(LocalCache));

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Deletes data for a specified key from the isolated storage.
        /// </summary>
        /// <param name="key">The key for which to delete the data.</param>
        public static void DeleteEntry(string key)
        {
            try
            {
                using (var store = GetIsolatedStorage())
                {
                    if (store.FileExists(key))
                    {
                        store.DeleteFile(key);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("Could not delete cache entry", ex);
                Debugger.Break();
            }
        }

        /// <summary>
        /// Reads data for a key and returns the data as a byte array.
        /// </summary>
        /// <param name="key">The key for which to read the data.</param>
        /// <returns>The data for the key as a byte array.</returns>
        public static byte[] ReadBytes(string key)
        {
            try
            {
                using (var store = GetIsolatedStorage())
                {
                    if (store.FileExists(key))
                    {
                        try
                        {
                            using (var reader = store.OpenFile(key, FileMode.Open, FileAccess.Read))
                            {
                                var bytes = new byte[reader.Length];
                                reader.Read(bytes, 0, (int)reader.Length);
                                return bytes;
                            }
                        }
                        catch (IsolatedStorageException ex)
                        {
                            Log.Error("Could not read from isolated storage file.", ex);
                            Debugger.Break();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("Could not read from isolated storage file.", ex);
                Debugger.Break();
            }
            return null;
        }

        /// <summary>
        /// Writes a byte array of data for a specified key.
        /// </summary>
        /// <param name="data">The data to write</param>
        /// <param name="key">The key for which to write the data</param>
        /// <returns><code>true</code> if writing was successful, <code>false</code> otherwise.</returns>
        public static bool WriteBytes(byte[] data, string key)
        {
            try
            {
                using (var store = GetIsolatedStorage())
                {
                    if (store.FileExists(key))
                    {
                        store.DeleteFile(key);
                    }
                    try
                    {
                        using (var writer = store.OpenFile(key, FileMode.Create, FileAccess.Write))
                        {
                            writer.Write(data, 0, data.Length);
                            writer.Flush();
                            return true;
                        }
                    }
                    catch (IsolatedStorageException ex)
                    {
                        Log.Error("Could not write to isolated storage file.", ex);
                        Debugger.Break();
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("Could not write to isolated storage file.", ex);
                Debugger.Break();
            }
            return false;
        }

        /// <summary>
        /// Method that gets the isolated storage file for the application.
        /// </summary>
        /// <returns>The isolated storage file.</returns>
        public static IsolatedStorageFile GetIsolatedStorage()
        {
            try
            {
                // Try to use application scoped isolated storage first
                return IsolatedStorageFile.GetUserStoreForApplication();
            }
            catch (IsolatedStorageException ex)
            {
                // Fallback to assembly-scoped isolated storage, when the application-scoped storage 
                // could not be initialized. This happens, e.g., when the application was not installed 
                // from the network drive, or the application was started in debug mode from VS.

                Log.Info("Could not initialize application scoped isolated storage file. Falling back to assembly scoped isolated storage file.", ex);
                return IsolatedStorageFile.GetUserStoreForAssembly();
            }
        }

        /// <summary>
        /// Reads an entire object from the local cache by its key.
        /// </summary>
        /// <param name="key">The key for which to read the object.</param>
        /// <param name="defaultFallbackObject">the object to use for writing in case that the reading is not successful</param>
        /// <returns>The object that was stored under the key.</returns>
        public static object ReadObject(string key, object defaultFallbackObject)
        {
            var formatter = new BinaryFormatter();

            var knownIssueData = LocalCache.ReadBytes(key);
            if (knownIssueData != null)
            {
                try
                {
                    using (var mem = new MemoryStream(knownIssueData))
                    {
                        var obj = formatter.Deserialize(mem);
                        return obj;
                    }
                }
                catch (Exception ex)
                {
                    Log.Info(string.Format("Error while deserializing data with key {0} from the local cache. Setting key to default.", key), ex);
                    try
                    {
                        LocalCache.WriteObject(key, defaultFallbackObject);
                    }
                    catch (Exception e)
                    {
                        Log.Error(string.Format("Error while setting data with key {0} to default.", key), e);
                    }
                    Debugger.Break();
                }
            }
            return null;
        }

        /// <summary>
        /// Writes an entire object under a specified key.
        /// </summary>
        /// <param name="key">The key under which to store the object.</param>
        /// <param name="value">The object which to store under the key</param>
        /// <returns><code>true</code> if the operation was successful, <code>false</code> otherwise.</returns>
        public static bool WriteObject(string key, object value)
        {
            var formatter = new BinaryFormatter();

            try
            {
                using (var target = new MemoryStream())
                {
                    formatter.Serialize(target, value);
                    LocalCache.WriteBytes(target.ToArray(), key);
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Error("Error while serializing and writing data to the local cache.", ex);
                Debugger.Break();
                return false;
            }
        }


        /// <summary>
        /// Method to update the known issues stored in the cache to contain the elementes stored in the local variable
        /// </summary>
        /// <param name="issueInfos">the issue infos to store in the cache</param>
        public static void UpdateKnownIssuesListInCache(Dictionary<int, IssueInfo> issueInfos)
        {
            // get copy of list of issues and store them
            var issueList = issueInfos.Values.ToList();
            var success = LocalCache.WriteObject(LocalCache.KnownIssues, issueList);
            if (success)
            {
                // if storing was successful, update last sync date
                Settings.Default.LastIssueSyncDate = DateTime.Now.Date;
                Settings.Default.Save();
            }
        }

        /// <summary>
        /// Method to get the list of known issues
        /// </summary>
        /// <returns>the list of known issues</returns>
        public static List<IssueInfo> GetKnownIssueListFromCache()
        {
            List<IssueInfo> knownIssueList;
            try
            {
                knownIssueList = LocalCache.ReadObject(LocalCache.KnownIssues, new List<IssueInfo>()) as List<IssueInfo> ?? new List<IssueInfo>();
            }
            catch (Exception exc)
            {
                Log.Error(
                    "Die Liste bekannter Issues konnte nicht deserialisiert werden, es wurde jetzt eine leere Liste verwendet und die Liste bekannter Issues im Cache zurückgesetzt.",
                    exc);
                knownIssueList = new List<IssueInfo>();
                LocalCache.WriteObject(LocalCache.KnownIssues, knownIssueList);
            }
            return knownIssueList;
        }


        #endregion
    }
}