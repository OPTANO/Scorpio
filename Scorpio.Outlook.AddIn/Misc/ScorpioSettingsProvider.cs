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
    using System;
    using System.Collections;
    using System.Collections.Specialized;
    using System.Configuration;
    using System.IO;
    using System.Text;
    using System.Xml;

    using log4net;

    /// <summary>
    /// Settings Provider for SCORPIO. This custom settings provider is needed, because the 
    /// settings are lost everytime there is an outlook update. That is because the default 
    /// setttings provider saves the settings in a plugin as well as outlook-version specific 
    /// way.
    /// See https://kikistidbits.blogspot.de/2010/10/save-your-settingssettings-to-known.html
    /// See https://msdn.microsoft.com/en-us/library/ms230624%28VS.90%29.aspx?f=255&amp;MSPPError=-2147217396
    /// </summary>
    public class ScorpioSettingsProvider : SettingsProvider
    {
        #region Static Fields

        /// <summary>
        /// The logger.
        /// </summary>
        private static readonly ILog Log = log4net.LogManager.GetLogger(typeof(ScorpioSettingsProvider));

        #endregion

        #region Public Properties

        #region Public properties

        /// <summary>
        /// Gets or sets the name of the application.
        /// </summary>
        public override string ApplicationName
        {
            get
            {
                return "SCORPIO";
            }
            set
            {
            }
        }

        #endregion

        #endregion

        #region Properties

        /// <summary>
        /// Gets the path to the settings file, which is be design not version dependant
        /// (in contrast to the default provider).
        /// The path is built as user's application data directory, then a subdirectory based
        /// on the application name, and the file name is always user.config similar to the
        /// default provider.
        /// </summary>
        private string GetSavingPath
        {
            get
            {
                var retVal = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + Path.DirectorySeparatorChar + this.ApplicationName
                             + Path.DirectorySeparatorChar + "SCORPIO.config";
                return retVal;
            }
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Gets the property value collection.
        /// </summary>
        /// <param name="context">The context</param>
        /// <param name="collection">The property collection</param>
        /// <returns>The settings property value collection</returns>
        public override SettingsPropertyValueCollection GetPropertyValues(SettingsContext context, SettingsPropertyCollection collection)
        {
            // Create new collection of values
            SettingsPropertyValueCollection values = new SettingsPropertyValueCollection();

            // Iterate through the settings to be retrieved (use their default values)
            foreach (SettingsProperty setting in collection)
            {
                SettingsPropertyValue value = new SettingsPropertyValue(setting);
                value.IsDirty = false;

                /*value.SerializedValue = setting.DefaultValue;
                value.PropertyValue = setting.DefaultValue;*/
                values.Add(value);
            }
            if (!File.Exists(this.GetSavingPath))
            {
                Log.Debug("Settings file does not exist (yet) - returning default values");
                return values;
            }
            try
            {
                using (var tr = new XmlTextReader(this.GetSavingPath))
                {
                    try
                    {
                        tr.ReadStartElement(this.ApplicationName);
                        foreach (SettingsPropertyValue value in values)
                        {
                            if (this.IsUserScoped(value.Property))
                            {
                                try
                                {
                                    tr.ReadStartElement(value.Name);
                                    value.SerializedValue = tr.ReadContentAsObject();
                                    value.Deserialized = false;
                                    tr.ReadEndElement();
                                }
                                catch (XmlException xe1)
                                {
                                    Log.Error("Failed to read value from settings file", xe1);
                                }
                            }
                        }
                        tr.ReadEndElement();
                    }
                    catch (XmlException xe2)
                    {
                        Log.Error("Failed to read section from settings file", xe2);
                    }
                }
            }
            catch (Exception e)
            {
                Log.Error("Failed to read settings file", e);
            }
            return values;
        }

        /// <summary>
        /// Here we just call the base class initializer
        /// </summary>
        /// <param name="name">The name</param>
        /// <param name="col">The name value collection</param>
        public override void Initialize(string name, NameValueCollection col)
        {
            base.Initialize(this.ApplicationName, col);
        }

        /// <summary>
        /// Sets the property values
        /// </summary>
        /// <param name="context">The context</param>
        /// <param name="collection">The settings property value collection</param>
        public override void SetPropertyValues(SettingsContext context, SettingsPropertyValueCollection collection)
        {
            var dir = Path.GetDirectoryName(this.GetSavingPath);
            if (!Directory.Exists(dir))
            {
                Log.Info("Settings directory does not exist, creating it");
                try
                {
                    Directory.CreateDirectory(dir);
                }
                catch (Exception fe)
                {
                    Log.ErrorFormat("Failed to create directory {0}", dir);
                    Log.Error("Exception", fe);
                }
            }
            try
            {
                using (var tw = new XmlTextWriter(this.GetSavingPath, Encoding.Unicode))
                {
                    tw.WriteStartDocument();
                    tw.WriteStartElement(this.ApplicationName);
                    foreach (SettingsPropertyValue propertyValue in collection)
                    {
                        if (this.IsUserScoped(propertyValue.Property) && propertyValue.SerializedValue != null)
                        {
                            tw.WriteStartElement(propertyValue.Name);
                            tw.WriteValue(propertyValue.SerializedValue);
                            tw.WriteEndElement();
                        }
                    }
                    tw.WriteEndElement();
                    tw.WriteEndDocument();
                }
            }
            catch (Exception e)
            {
                Log.Error("Unable to save settings", e);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Helper method: walks the "attribute bag" for a given property
        /// to determine if it is user-scoped or not.
        /// Note that this provider does not enforce other rules, such as
        /// - unknown attributes
        /// - improper attribute combinations (e.g. both user and app - this implementation
        /// would say true for user-scoped regardless of existence of app-scoped)
        /// </summary>
        /// <param name="prop">The property for which to determine if it is user scoped.</param>
        /// <returns><code>true</code> if the property is user scoped, <code>false</code> otherwise.</returns>
        private bool IsUserScoped(SettingsProperty prop)
        {
            foreach (DictionaryEntry d in prop.Attributes)
            {
                var a = (Attribute)d.Value;
                if (a is UserScopedSettingAttribute)
                {
                    return true;
                }
            }
            return false;
        }

        #endregion
    }
}