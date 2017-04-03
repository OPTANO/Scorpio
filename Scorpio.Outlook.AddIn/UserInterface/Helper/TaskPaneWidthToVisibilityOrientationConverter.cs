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

namespace Scorpio.Outlook.AddIn.UserInterface.Helper
{
    using System;
    using System.Globalization;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Data;

    /// <summary>
    /// Value converter for WPF bindings which converts an integer to a visibility.
    /// </summary>
    public class TaskPaneWidthToVisibilityOrientationConverter : IValueConverter
    {
        #region Public Methods and Operators

        /// <summary>
        /// The convert method. It takes an integer value, and compares it to a parameter. If the value is greater than the parameter, the converter returns Visible, otherwise, the converter return Collapsed.
        /// </summary>
        /// <param name="value">The integer value to be converted</param>
        /// <param name="targetType">The target type of the conversion</param>
        /// <param name="parameter">The optional parameter for the conversion. This should be a number representing the threshhold for visibility.</param>
        /// <param name="culture">The culture</param>
        /// <returns>Visible if the value is larger than the threshhold-parameter, Collapsed otherwise.</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var threshhold = 200.0;
            if (parameter != null)
            {
                var convertedParameter = parameter as string;
                if (convertedParameter != null)
                {
                    double.TryParse(convertedParameter, out threshhold);
                }
            }

            var convertedValue = value as double?;

            if (!convertedValue.HasValue)
            {
                if (targetType == typeof(Visibility))
                {
                    return Visibility.Collapsed;
                }
                if (targetType == typeof(Orientation))
                {
                    return Orientation.Vertical;
                }
                return null;
            }
            if (targetType == typeof(Visibility))
            {
                return convertedValue < threshhold ? Visibility.Collapsed : Visibility.Visible;
            }
            if (targetType == typeof(Orientation))
            {
                return convertedValue < threshhold ? Orientation.Vertical : Orientation.Horizontal;
            }
            return null;
        }

        /// <summary>
        /// No conversion back to int.
        /// </summary>
        /// <param name="value">The value</param>
        /// <param name="targetType">The target type</param>
        /// <param name="parameter">The parameter</param>
        /// <param name="culture">The culture</param>
        /// <returns>Nothing, because the method is not implemented.</returns>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}