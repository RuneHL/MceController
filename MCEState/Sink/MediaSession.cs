/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 */
using ehiProxy;
using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Text;
//using Microsoft.Ehome.Epg;
using Microsoft.MediaCenter.TV.Scheduling;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.DirectShow.Metadata;
using System.Globalization;
using VmcController.Services;
//using VmcController.Services;

namespace VmcController.MceState
{
    [Guid("92f0245d-7608-4bab-b0ab-8c596a617656"), ComVisible(true)]
    public class MediaSession : IMediaStatusSession
    {
        #region Member Variables
        /// <summary>
        /// Our session Id
        /// </summary>
        private int m_sessionId;
        /// <summary>
        /// Previous tag (used for removing duplicates)
        /// </summary>
        private MEDIASTATUSPROPERTYTAG m_prevTag;
        /// <summary>
        /// Previous property value (used for removing duplicates)
        /// </summary>
        private object m_prevProp;
        #endregion

        #region Constructor
        /// <summary>
        /// Session constructor
        /// </summary>
        /// <param name="id">session id</param>
        public MediaSession(int id)
        {
            this.m_sessionId = id;
            if (MediaSink.SocketServer != null)
                MediaSink.SocketServer.SendMessage(string.Format(CultureInfo.InvariantCulture, "StartSession={0}\r\n", m_sessionId) );
            else
                Trace.TraceWarning("SocketServer reference is null");
        }
        #endregion

        #region IMediaStatusSession Implementation Methods
        /// <summary>
        /// Event handler for media status changes
        /// </summary>
        /// <param name="Tags">Array of tags</param>
        /// <param name="Properties">Array of property values</param>
        public void MediaStatusChange(MEDIASTATUSPROPERTYTAG[] tags, object[] properties)
        {
            //  Check to see if the socket server is valid
            if (MediaSink.SocketServer == null) return;

            try
            {
                for (int i = 0; i < tags.Length; i++)
                {
                    MEDIASTATUSPROPERTYTAG tag = tags[i];
                    object property = properties[i];
                    //  Remove duplicate notifications
                    if ((tag == m_prevTag) && (property.ToString() == m_prevProp.ToString()))
                        return;
                    else
                    {
                        m_prevTag = tag;
                        m_prevProp = property;
                    }

                    Trace.TraceInformation("{0}={1}", tag, property);

                    //  Announce the status change to all listeners
                    MediaSink.SocketServer.SendMessage(string.Format(CultureInfo.InvariantCulture, "{0}={1}\r\n", tag, property));

                    //  Update the media state information
                    MediaState.UpdateState(tag, property);

                    //  Check for MediaName status changes and get additional EPG data
                    //if (tag == MEDIASTATUSPROPERTYTAG.MediaName && MediaState.MediaMode == MEDIASTATUSPROPERTYTAG.TVTuner)
                    //{
                        //  Look up epg info in a background worker thread
                    //    ThreadPool.QueueUserWorkItem(new WaitCallback(LookupEpgInfo), property);
                    //}
                    //else if (tag == MEDIASTATUSPROPERTYTAG.MediaName && MediaState.MediaMode == MEDIASTATUSPROPERTYTAG.PVR)
                    //{
                    //    //  Get pvr media info in a background worker thread
                    //    ThreadPool.QueueUserWorkItem(new WaitCallback(RetrieveDvrmsInfo));
                    //}
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("Exception: {0}", ex.ToString());
            }
        }

        /// <summary>
        /// Close the media session
        /// </summary>
        public void Close()
        {
            Trace.TraceInformation("Closing media session #{0}", m_sessionId);
            if (MediaSink.SocketServer != null)
                MediaSink.SocketServer.SendMessage(string.Format(CultureInfo.InvariantCulture, "EndSession={0}\r\n", m_sessionId));
        }
        #endregion

        #region Helper Methods
        /// <summary>
        /// Sends the status change event to all clients.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="value">The value.</param>
        private static void SendCustomStatus(string key, object value)
        {
            Trace.TraceInformation("{0}={1}", key, value);
            MediaState.MetaData[key] = value;
            if (MediaSink.SocketServer != null)
                MediaSink.SocketServer.SendMessage(string.Format(CultureInfo.InvariantCulture, "{0}={1}\r\n", key, value));
        }

        /// <summary>
        /// Lookup the epg info for the currently playing show to find additonal EPG meta data
        /// </summary>
        /// <param name="stateInfo">The state info.</param>
        /// 
        /*
        private static void LookupEpgInfo(Object parameter)
        {
            Trace.TraceInformation("LookupEpgInfo() Start");
            try
            {
                Guide guide = Guide.GuideInstance as Guide;
                string mediaName = parameter.ToString();

                //  Iterate through the channels checking for a media name match
                foreach (Channel channel in guide.Channels)
                {
                    ScheduleEntry scheduleEntry = channel.ShowAt(DateTime.Now.ToUniversalTime());
                    if (scheduleEntry.Program.Title.Equals(mediaName))
                    {
                        SendCustomStatus("EpisodeTitle", scheduleEntry.Program.EpisodeTitle);
                        SendCustomStatus("Description", scheduleEntry.Program.Description);
                        SendCustomStatus("ChannelNumber", scheduleEntry.Channel.Number);
                        SendCustomStatus("ChannelName", scheduleEntry.Channel);
                        SendCustomStatus("ChannelCallsign", scheduleEntry.Channel.DefaultService.CallSign);
                        SendCustomStatus("HDTV", scheduleEntry.HDTV);
                        return;
                    }
                }
                Trace.TraceWarning("No Epg information found");
            }
            catch (Exception ex)
            {
                Trace.TraceError("Exception: {0}", ex.ToString());
            }
            finally
            {
                Trace.Unindent();
                Trace.TraceInformation("LookupEpgInfo() Exit");
            }
        }
        */

        /// <summary>
        /// Lookups the PVR info for the currently playing show to find additional EPG meta data
        /// </summary>
        /// <param name="parameter">The parameter.</param>
        private static void RetrieveDvrmsInfo(Object parameter)
        {
            int pid = 0;
            int mySession;
            string filename = null;

            Trace.TraceInformation("RetrieveDvrmsInfo() Start");
            Trace.Indent();
            try
            {
                //  Get the session ID for the current process
                using (Process currentProcess = Process.GetCurrentProcess())
                    mySession = currentProcess.SessionId;

                //  Find the matching ehshell process in the same session
                foreach (Process proc in Process.GetProcessesByName("ehshell"))
                {
                    if (proc.SessionId == mySession) pid = proc.Id;
                    proc.Dispose();
                }

                //  Check we found a process, if not, abort
                if (pid == 0)
                {
                    Trace.TraceWarning("No ehshell process found");
                    return;
                }

                //  Get the list of open files by the ehshell process
                using (IEnumerator<FileSystemInfo> openFiles = DetectOpenFiles.GetOpenFilesEnumerator(pid))
                {
                    //  Iterate until we find a .dvr-ms file
                    while (openFiles.MoveNext())
                    {
                        Trace.TraceInformation("Openfile detected: {0}", openFiles.Current.Name);
                        if (openFiles.Current.Extension.ToLower(CultureInfo.InvariantCulture) == ".dvr-ms")
                        {
                            filename = openFiles.Current.FullName;
                            break;
                        }
                    }
                }

                //  Check we found a DVR-MS file
                if (string.IsNullOrEmpty(filename))
                {
                    Trace.TraceWarning("No DVR-MS open file found");
                    return;
                }

                //  Extract meta data and send to clients
                using (DvrmsMetadata dvrmsMetadata = new DvrmsMetadata(filename))
                {
                    Dictionary<string, MetadataItem> metaData = dvrmsMetadata.GetAttributes();
                    SendCustomStatus("EpisodeTitle", metaData["WM/SubTitle"].Value);
                    SendCustomStatus("Description", metaData["WM/SubTitleDescription"].Value);
                    SendCustomStatus("ChannelNumber", metaData["WM/MediaOriginalChannel"].Value);
                    SendCustomStatus("ChannelName", metaData["WM/MediaStationName"].Value);
                    SendCustomStatus("ChannelCallsign", metaData["WM/MediaStationCallSign"].Value);
                    SendCustomStatus("ParentalAdvisoryRating", metaData["WM/ParentalRating"].Value);
                    SendCustomStatus("HDTV", metaData["WM/WMRVHDContent"].Value);
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("Exception: {0}", ex.ToString());
            }
            finally
            {
                Trace.Unindent();
                Trace.TraceInformation("RetrieveDvrmsInfo() Exit");
            }
        }
        #endregion
    }
}
