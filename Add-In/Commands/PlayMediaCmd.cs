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
using System.Collections.Generic;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;
using ehiProxy;

namespace VmcController.AddIn.Commands {
	/// <summary>
	/// Summary description for MsgBox command.
	/// </summary>
	public class PlayMediaCmd : ICommand {
		private MediaType m_mediaType;
		private bool m_appendToQueue;

		public PlayMediaCmd(MediaType mediaType, bool appendToQueue) {
			m_mediaType = mediaType;
			m_appendToQueue = appendToQueue;
		}

		#region ICommand Members

		/// <summary>
		/// Shows the syntax.
		/// </summary>
		/// <returns></returns>
		public string ShowSyntax() {
			return "<" + m_mediaType.ToString() + " parameters>";
		}

		/// <summary>
		/// Executes the specified param.
		/// </summary>
		/// <param name="param">The param.</param>
		/// <param name="result">The result.</param>
		/// <returns></returns>
		public OpResult Execute(string param) {
			OpResult opResult = new OpResult();
			try {
				switch (m_mediaType) {
					case MediaType.Dvd:
						param = param.Replace('\\', '/');
						break;
					case MediaType.Audio:
					case MediaType.Video:
					case MediaType.Dvr:
						param = param.Replace(@"\", @"\\");
						break;
						//case MediaType.TV:
						//    //  Check for a channel number and convert to a callsign
						//    int channelNum;
						//    if (int.TryParse(param, out channelNum))
						//    {
						//        Channel channel = Guide.GuideInstance.GetChannelByNumber(channelNum);
						//        if (channel != null) param = channel.DefaultService.CallSign;
						//    }
						//    break;
				}
				if (AddInHost.Current.MediaCenterEnvironment.PlayMedia(m_mediaType, param, m_appendToQueue)) {
					opResult.StatusCode = OpStatusCode.Success;
				}
				else {
					opResult.StatusCode = OpStatusCode.BadRequest;
				}
			}
			catch (Exception ex) {
				opResult.StatusCode = OpStatusCode.Exception;
				opResult.StatusText = ex.Message;
			}
			return opResult;
		}

		#endregion
	}
}