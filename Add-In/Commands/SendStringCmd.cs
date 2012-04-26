/*
 * This module is build on top of on J.Bradshaw's vmcController
 *
 * Implements string sending
 * 
 * Copyright (c) 2009 Anthony Jones
 * 
 * Portions copyright (c) 2007 Jonathan Bradshaw
 * 
 * This software code module is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely.
 * 
 * 1. The origin of this software must not be misrepresented; you must not claim that you wrote 
 *    the original software. If you use this software in a product, an acknowledgment in the 
 *    product documentation would be appreciated but is not required.
 * 2. Altered source versions must be plainly marked as such, and must not be misrepresented as
 *    being the original software.
 * 3. This notice may not be removed or altered from any source distribution.
 * 
 ********************************************************************************
 * Changes:
 * 2008 Oct 6 - ajones@pobox.com - Added SendStringCmd.cs to send a string (implementing the 'type' command)
 * 2009 Sep 1 - ajones@pobox.com - Added support for shifted keys (allowing !@#$ etc)
 * 
 */
using System;
using System.Runtime.InteropServices;

namespace VmcController.AddIn.Commands
{
    /// <summary>
    /// Summary description for SendString command.
    /// </summary>
    public class SendStringCmd : ICommand
    {
        public bool debug = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="SendString"/> class.
        /// </summary>
        /// <param name="vk">The vk.</param>
        public SendStringCmd()
        {
        }

        public SendStringCmd(bool debug)
		{
			this.debug = debug;
		}

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "<text> - sends \"<text>\" to application";
        }

        /// <summary>
        /// Sends the specified string to the application as keystrokes.
        /// </summary>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            OpResult opResult = new OpResult();
            ushort key = 0;
            try
            {
                param = param.ToUpper();
                foreach (char c in param.ToCharArray(0, param.Length))
                {
                    key = (ushort)NativeMethods.VkKeyScan(c);
                    bool shift = ((NativeMethods.VkKeyScan(c) & 0x0100) > 0);

                    SendKeyCmd skc = new SendKeyCmd(key, shift, false, false);
                    opResult = skc.Execute("");
                    if (debug)
                    {
                        MsgBoxCmd mbc = new MsgBoxCmd();
                        string s = String.Format("\"Just sent a character!\" \"Character sent: '{0}', vk: {1}\" 1", c, key);
                        mbc.Execute(s);
                    }
                    if (opResult.StatusCode != OpStatusCode.Success) break;
                }
                if (debug)
                {
                    MsgBoxCmd mbc = new MsgBoxCmd();
                    mbc.Execute("\"Length: " + param.Length + "\" \"Sent: '" + param + "'\" 5");
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

        #region Win32 Helpers
        static class NativeMethods
        {
            public const int INPUT_KEYBOARD = 1;
            public const uint KEY_UP = 0x0002;

            [StructLayout(LayoutKind.Sequential)]
            public struct KEYBDINPUT
            {
                public ushort wVk;
                public ushort wScan;
                public uint dwFlags;
                public uint time;
                public IntPtr dwExtraInfo;
            };

            [StructLayout(LayoutKind.Explicit, Size = 28)]
            public struct INPUT
            {
                [FieldOffset(0)]
                public uint type;
                [FieldOffset(4)]
                public KEYBDINPUT ki;
            };

            [DllImport("user32.dll")]
            public static extern short VkKeyScan(char ch);

            [DllImport("user32.dll")]
            public static extern uint SendInput(uint nInputs, ref INPUT pInputs, int cbSize);

            // Activate an application window.
            [DllImport("USER32.DLL")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
        }
        #endregion
    }
}
