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
using System.Text.RegularExpressions;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for MsgBox command.
	/// </summary>
	public class SendKeyCmd : ICommand
    {
        private ushort vk = 0;
		public bool			Shift = false;
		public bool			Ctrl = false;
		public bool			Alt = false;
		public bool			Win = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="SendKey"/> class.
        /// </summary>
        /// <param name="vk">The vk.</param>
        public SendKeyCmd(ushort vk)
        {
            this.vk = vk;
            this.Shift = false;
            this.Ctrl = false;
            this.Alt = false;
            this.Win = false;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SendKey"/> class.
        /// </summary>
        /// <param name="vk">The vk.</param>
        /// <param name="Shift">if set to <c>true</c> [shift].</param>
        /// <param name="Ctrl">if set to <c>true</c> [CTRL].</param>
        /// <param name="Alt">if set to <c>true</c> [alt].</param>
        public SendKeyCmd(ushort vk, bool Shift, bool Ctrl, bool Alt)
		{
			this.vk = vk;
			this.Shift = Shift;
			this.Ctrl = Ctrl;
			this.Alt = Alt;
			this.Win = false;
		}

        /// <summary>
        /// Initializes a new instance of the <see cref="SendKey"/> class.
        /// </summary>
        /// <param name="vk">The vk.</param>
        /// <param name="Shift">if set to <c>true</c> [shift].</param>
        /// <param name="Ctrl">if set to <c>true</c> [CTRL].</param>
        /// <param name="Alt">if set to <c>true</c> [alt].</param>
        /// <param name="Win">if set to <c>true</c> [win].</param>
        public SendKeyCmd(ushort vk, bool Shift, bool Ctrl, bool Alt, bool Win) 
		{
			this.vk = vk;
			this.Shift = Shift;
			this.Ctrl = Ctrl;
			this.Alt = Alt;
			this.Win = Win;
		}

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "- sends key to application";
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
            NativeMethods.INPUT structInput;
            structInput = new NativeMethods.INPUT();
            structInput.type = NativeMethods.INPUT_KEYBOARD;
            structInput.ki.wScan = 0;
            structInput.ki.time = 0;
            structInput.ki.dwFlags = 0;
            structInput.ki.dwExtraInfo = IntPtr.Zero;

            try
            {
                while (NativeMethods.SetForegroundWindow(GetWindowHandle()) == false)
                {}
                if (NativeMethods.SetForegroundWindow(GetWindowHandle()) == false)
                {
                    opResult.StatusCode = OpStatusCode.BadRequest;
                    return opResult;
                }

                if (Shift)
                {
                    structInput.ki.wVk = 0x10;
                    NativeMethods.SendInput(1, ref structInput, Marshal.SizeOf(structInput));
                }
                if (Ctrl)
                {
                    structInput.ki.wVk = 0x11;
                    NativeMethods.SendInput(1, ref structInput, Marshal.SizeOf(structInput));
                }
                if (Alt)
                {
                    structInput.ki.wVk = 0x12;
                    NativeMethods.SendInput(1, ref structInput, Marshal.SizeOf(structInput));
                }
                if (Win)
                {
                    structInput.ki.wVk = 0x5b;
                    NativeMethods.SendInput(1, ref structInput, Marshal.SizeOf(structInput));
                }
                structInput.ki.wVk = vk;
                // Key down the actual key-code
                NativeMethods.SendInput(1, ref structInput, Marshal.SizeOf(structInput));
                // Key up the actual key-code
                structInput.ki.dwFlags = NativeMethods.KEY_UP;
                NativeMethods.SendInput(1, ref structInput, Marshal.SizeOf(structInput));
                if (Shift)
                {
                    structInput.ki.wVk = 0x10;
                    NativeMethods.SendInput(1, ref structInput, Marshal.SizeOf(structInput));
                }
                if (Ctrl)
                {
                    structInput.ki.wVk = 0x11;
                    NativeMethods.SendInput(1, ref structInput, Marshal.SizeOf(structInput));
                }
                if (Alt)
                {
                    structInput.ki.wVk = 0x12;
                    NativeMethods.SendInput(1, ref structInput, Marshal.SizeOf(structInput));
                }
                if (Win)
                {
                    structInput.ki.wVk = 0x5b;
                    NativeMethods.SendInput(1, ref structInput, Marshal.SizeOf(structInput));
                }
                opResult.StatusCode = OpStatusCode.Success;
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

            [StructLayout(LayoutKind.Sequential)]
            public struct MOUSEINPUT
            {
                public int dx;
                public int dy;
                public uint mouseData;
                public uint dwFlags;
                public uint time;
                public IntPtr dwExtraInfo;
            }

            [StructLayout(LayoutKind.Explicit)]
            public struct INPUT
            {
                [FieldOffset(0)]
                public uint type;
                [FieldOffset(8)]
                public MOUSEINPUT mi;
                [FieldOffset(8)]
                public KEYBDINPUT ki;
            };

            [DllImport("user32.dll")]
            public static extern uint SendInput(uint nInputs, ref INPUT pInputs, int cbSize);

            // Activate an application window.
            [DllImport("USER32.DLL")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
        }
        #endregion
    }
}
