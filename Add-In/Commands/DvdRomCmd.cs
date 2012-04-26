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
using System.Text;
using System.Runtime.InteropServices;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for DvdRom command.
	/// </summary>
	public class DvdRomCmd: ICommand
	{
        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "<open|close> - Opens and closes the default dvd drive door";
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            OpResult opResult = new OpResult(OpStatusCode.BadRequest);
            StringBuilder rt = new StringBuilder();

            try
            {
                if (param.Equals("open", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (mciSendString("set CDAudio door open", rt, 127, IntPtr.Zero) == 0)
                        opResult.StatusCode = OpStatusCode.Ok;
                }
                else
                {
                    if (mciSendString("set CDAudio door closed", rt, 127, IntPtr.Zero) == 0)
                        opResult.StatusCode = OpStatusCode.Ok;
                }
            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = ex.Message;
            }
            return opResult;
        }

        #endregion

        #region External Imports
        [DllImport("winmm.dll")]
        static extern Int32 mciSendString(String command,
           StringBuilder buffer, Int32 bufferSize, IntPtr hwndCallback);
        #endregion
    }
}
