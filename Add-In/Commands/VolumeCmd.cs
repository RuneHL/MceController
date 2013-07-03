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

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for Volume command.
	/// </summary>
	public class Volume : ICommand
	{
        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "<0-50|Up|Down|Mute|UnMute|Get>";
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            OpResult opResult = new OpResult(OpStatusCode.Success);
            try
            {
                if (param.Equals("Up", StringComparison.InvariantCultureIgnoreCase))
                    AddInHost.Current.MediaCenterEnvironment.AudioMixer.VolumeUp();
                else if (param.Equals("Down", StringComparison.InvariantCultureIgnoreCase))
                    AddInHost.Current.MediaCenterEnvironment.AudioMixer.VolumeDown();
                else if (param.Equals("Mute", StringComparison.InvariantCultureIgnoreCase))
                    AddInHost.Current.MediaCenterEnvironment.AudioMixer.Mute = true;
                else if (param.Equals("UnMute", StringComparison.InvariantCultureIgnoreCase))
                    AddInHost.Current.MediaCenterEnvironment.AudioMixer.Mute = false;
                else if (param.Equals("Get", StringComparison.InvariantCultureIgnoreCase))
                {
                    opResult.StatusCode = OpStatusCode.Ok;
                    opResult.AppendFormat("volume={0}", (int)(AddInHost.Current.MediaCenterEnvironment.AudioMixer.Volume / 1310.7));
                }
                else
                {
                    int desiredLevel = int.Parse(param);
                    if (desiredLevel > 50 || desiredLevel < 0)
                    {
                        opResult.StatusCode = OpStatusCode.BadRequest;
                        return opResult;
                    }

                    int volume = (int)(AddInHost.Current.MediaCenterEnvironment.AudioMixer.Volume / 1310.7);
                    for (int level = volume; level > desiredLevel; level--)
                        AddInHost.Current.MediaCenterEnvironment.AudioMixer.VolumeDown();

                    for (int level = volume; level < desiredLevel; level++)
                        AddInHost.Current.MediaCenterEnvironment.AudioMixer.VolumeUp();
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
    }
}
