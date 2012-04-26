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
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
//using Microsoft.Ehome.Epg;
using ehiProxy;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for MsgBox command.
	/// </summary>
	public class EpgCmd : ICommand
	{
        private static Regex m_regex = new Regex(@"(?<channel>all|[\d,]+)");
        private EpgResultsType m_resultsType;

        /// <summary>
        /// Initializes a new instance of the <see cref="EpgCmd"/> class.
        /// </summary>
        /// <param name="resultsLevel">The results level.</param>
        public EpgCmd(EpgResultsType resultsType)
        {
            m_resultsType = resultsType;
        }

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "<all|channel,channel...>";
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
                Match match = m_regex.Match(param);
                if (!match.Success)
                {
                    opResult.StatusCode = OpStatusCode.BadRequest;
                    return opResult;
                }
                
                /*
                Guide guide = Guide.GuideInstance as Guide;
                if (match.Groups["channel"].Value.Equals("all", StringComparison.InvariantCultureIgnoreCase))
                {
                    foreach (Channel channel in guide.Channels)
                        GetChannelEpg(channel, opResult);
                }
                else
                {
                    foreach (string numStr in match.Groups["channel"].Value.Split(','))
                        GetChannelEpg(guide.GetChannelByNumberString(numStr), opResult);
                }
                 * */

                opResult.StatusCode = OpStatusCode.Ok;
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
        /// Gets the channel epg.
        /// </summary>
        /// <param name="channel">The channel.</param>
        /// <returns></returns>
        /*
        private void GetChannelEpg(Channel channel, OpResult opResult)
        {
            ScheduleEntry entry;

            if (channel == null)
                return;

            switch (m_resultsType)
            {
                case EpgResultsType.ChannelName:
                    opResult.AppendFormat("{0}={0} {1}", channel.Number, channel.ToString());
                    break;

                case EpgResultsType.OnNow:
                    entry = channel.ShowAt(DateTime.Now.ToUniversalTime());
                    opResult.AppendFormat("{0}={1} ({2}-{3})", 
                        channel.Number, 
                        entry.Program.ToString(),
                        entry.StartTime.ToLocalTime().ToString("g"),
                        entry.EndTime.ToLocalTime().ToShortTimeString());
                    break;

                case EpgResultsType.Details:
                    entry = channel.ShowAt(DateTime.Now.ToUniversalTime());
                    opResult.AppendFormat("Title={0}", entry.Program.Title);
                    opResult.AppendFormat("EpisodeTitle={0}", entry.Program.EpisodeTitle);
                    opResult.AppendFormat("Description={0}", entry.Program.Description);
                    opResult.AppendFormat("ChannelNumber={0}", channel.Number);
                    opResult.AppendFormat("ChannelName={0}", channel.ToString());
                    opResult.AppendFormat("ChannelCallsign={0}", channel.DefaultService.CallSign);
                    opResult.AppendFormat("StartTime={0}", entry.StartTime.ToLocalTime().ToString("s"));
                    opResult.AppendFormat("EndTime={0}", entry.EndTime.ToLocalTime().ToString("s"));
                    opResult.AppendFormat("Duration={0}", entry.Duration.ToString());
                    opResult.AppendFormat("HDTV={0}", entry.HDTV);
                    break;
            }
        }
        */
    }

    public enum EpgResultsType
    {
        ChannelName,
        OnNow,
        Details
    }
}
