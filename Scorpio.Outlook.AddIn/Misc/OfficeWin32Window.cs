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
    using System.Runtime.InteropServices;
    using System.Windows.Interop;

    /// <summary>
    /// See https://stackoverflow.com/questions/12733974/how-to-set-the-window-owner-to-outlook-window
    /// This class retrieves the IWin32Window from the current active Office window.
    /// This could be used to set the parent for Windows Forms and MessageBoxes.
    /// </summary>
    /// <example>
    /// OfficeWin32Window parentWindow = new OfficeWin32Window (ThisAddIn.OutlookApplication.ActiveWindow ());   
    /// MessageBox.Show (parentWindow, "This MessageBox doesn't go behind Outlook !!!", "Attention !", MessageBoxButtons.Ok , MessageBoxIcon.Question );
    /// </example>
    public class OfficeWin32Window : IWin32Window
    {
        #region Fields

        /// <summary>
        /// This holds the window handle for the found Window.
        /// </summary>
        private IntPtr _windowHandle = IntPtr.Zero;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OfficeWin32Window"/> class. Could be used to get the parent IWin32Window for Windows.Forms and MessageBoxes.
        /// </summary>
        /// <param name="windowObject">The current WindowObject.</param>
        public OfficeWin32Window(object windowObject)
        {
            string caption =
                windowObject.GetType().InvokeMember("Caption", System.Reflection.BindingFlags.GetProperty, null, windowObject, null).ToString();

            // try to get the HWND ptr from the windowObject / could be an Inspector window or an explorer window
            this._windowHandle = FindWindow("rctrl_renwnd32\0", caption);
        }

        #endregion

        #region Public Properties

        #region Public properties

        /// <summary>
        /// Gets the <b>Handle</b> of the Outlook WindowObject.
        /// </summary>
        public IntPtr Handle
        {
            get
            {
                return this._windowHandle;
            }
        }

        #endregion

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The <b>FindWindow</b> method finds a window by it's classname and caption.
        /// </summary>
        /// <param name="className">The classname of the window (use Spy++)</param>
        /// <param name="windowName">The Caption of the window.</param>
        /// <returns>Returns a valid window handle or 0.</returns>
        [DllImport("user32")]
        public static extern IntPtr FindWindow(string className, string windowName);

        #endregion
    }
}