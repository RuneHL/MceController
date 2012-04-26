/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 * This software is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely, subject to the following restrictions:
 * 
 * 1. The origin of this software must not be misrepresented; you must not claim that you wrote 
 *    the original software. If you use this software in a product, an acknowledgment in the 
 *    product documentation would be appreciated but is not required.
 * 2. Altered source versions must be plainly marked as such, and must not be misrepresented as
 *    being the original software.
 * 3. This notice may not be removed or altered from any source distribution.
 * 
 */
using System;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for MsgBox command.
	/// </summary>
	public class SysCommand : ICommand
    {
        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_MINIMIZE = 0xF020;
        public const int SC_MAXIMIZE = 0xF030;
        public const int SC_CLOSE = 0xF060;
        public const int SC_RESTORE = 0xF120;

		private int wParam;
		private int lParam;

        /// <summary>
        /// Initializes a new instance of the <see cref="SysCommand"/> class.
        /// </summary>
        /// <param name="wParam">The w param.</param>
        public SysCommand(int wParam)
        {
            this.wParam = wParam;
            this.lParam = 0;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SysCommand"/> class.
        /// </summary>
        /// <param name="wParam">The w param.</param>
        /// <param name="lParam">The l param.</param>
        public SysCommand(int wParam, int lParam)
		{
			this.wParam = wParam;
			this.lParam = lParam;
		}

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "- sends command to Media Center Window";
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            OpResult opResult = new OpResult();
            try
            {
                if (NativeMethods.SendMessage(GetWindowHandle(), WM_SYSCOMMAND, wParam, lParam) == 0)
                    opResult.StatusCode = OpStatusCode.Success;
                else
                    opResult.StatusCode = OpStatusCode.BadRequest;
            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = ex.Message;
            }
            return opResult;
        }

        #endregion

        /// <summary>
        /// Helper method to get the window handle for the main MCE shell in our session
        /// </summary>
        /// <returns></returns>
        private IntPtr GetWindowHandle()
        {
            IntPtr hwnd = IntPtr.Zero;
            int mySession;

            using (Process currentProcess = Process.GetCurrentProcess())
                mySession = currentProcess.SessionId;

            foreach (Process p in Process.GetProcessesByName("ehshell"))
            {
                if (p.SessionId == mySession)
                    hwnd = p.MainWindowHandle;
                p.Dispose();
            }
            return hwnd;
        }

        #region External SendMessage Import
        internal static class NativeMethods
        {
            [DllImport("user32.dll")]
            internal static extern int SendMessage(
                IntPtr hWnd, // handle to destination window
                uint Msg, // message
                int wParam, // first message parameter
                int lParam // second message parameter
            );
        }
        #endregion
    }
}
