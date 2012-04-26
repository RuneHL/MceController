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
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for PlayRate command.
	/// </summary>
	public class PlayRateCmd : ICommand
	{
        private bool m_set = true;
        private bool m_state = false;

        public PlayRateCmd(bool bSet)
        {
            m_set = bSet;
            m_state = false;
        }
        public PlayRateCmd(bool bSet, bool bState)
        {
            m_set = bSet;
            m_state = bState;
        }
        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            String s;
            if (!m_state)
            {
                StringBuilder sb = new StringBuilder("");
                foreach (string value in Enum.GetNames(typeof(PlayRateEnum)))
                    sb.AppendFormat("{0}|", value);
                sb.Remove(sb.Length - 1, 1);

                if (m_set) s = "<" + sb.ToString() + "> - sets the play rate";
                else s = "- returns the play rate (one of " + sb.ToString() + ")";
            }
            else s = "- returns the current play state";

            return s;
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
                if (AddInHost.Current.MediaCenterEnvironment.MediaExperience == null)
                {
                    if (m_set)
                    {
                        opResult.StatusCode = OpStatusCode.Ok;
                        opResult.AppendFormat("No media playing");
                    }
                    else
                    {
                        opResult.StatusCode = OpStatusCode.BadRequest;
                        opResult.AppendFormat("No media playing");
                    }
                }
                else if (m_set)
                {
                    PlayRateEnum playRate = (PlayRateEnum)Enum.Parse(typeof(PlayRateEnum), param, true);
                    if (playRate == PlayRateEnum.SkipForward)
                        AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.SkipForward();
                    else if (playRate == PlayRateEnum.SkipBack)
                        AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.SkipBack();
                    else
                        AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.PlayRate = (Single)playRate;
                }
                else if (!m_state)
                {
                    int rate = (int)AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.PlayRate;
                    opResult.AppendFormat("PlayRate={0}", Enum.GetNames(typeof(PlayRateEnum))[rate]);
                }
                else
                {
                    PlayState state = AddInHost.Current.MediaCenterEnvironment.MediaExperience.Transport.PlayState;
                    opResult.AppendFormat("PlayState={0}", Enum.GetName(typeof(PlayState), state));
                }
                opResult.StatusCode = OpStatusCode.Success;
            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.AppendFormat(ex.Message);
            }
            return opResult;
        }

        #endregion
    }

    enum PlayRateEnum
    {
        Stop,
        Pause,
        Play,
        FF1,
        FF2,
        FF3,
        Rewind1,
        Rewind2,
        Rewind3,
        SlowMotion1,
        SlowMotion2,
        SlowMotion3,
        SkipForward,
        SkipBack
    }

}
