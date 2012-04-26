/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 */
using System;
using System.IO;
using System.Diagnostics;
using System.Globalization;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using Microsoft.MediaCenter.TV.Scheduling;
using Microsoft.Win32;
using System.Text;

namespace VmcController.MceState
{
	/// <summary>
	/// Summary description for MSASSink.
	/// </summary>
    [Guid("392a06a8-064f-4d1d-af73-55cb9044b5d0"), ComVisible(true)]
	public sealed class MediaSink : IMediaStatusSink
    {
        #region Member Variables
        /// <summary>
        /// The base IP port number, this is incremented for each client
        /// </summary>
        private const int BASE_PORT = 40400;

        /// <summary>
        /// The port number for the socket server to bind to
        /// </summary>
        private static int m_portNumber;

        /// <summary>
        /// Count of media sessions
        /// </summary>
        private static int m_sessionCount;

        /// <summary>
        /// The TCP Socket Server
        /// </summary>
        private static TcpSocketServer m_socketServer;

        /// <summary>
        /// Media Center Recording Event Schedule
        /// </summary>
        private static EventSchedule m_eventSchedule;

        /// <summary>
        /// Keyboard hook used to send out keyboard events
        /// </summary>
        private static KeyboardHook m_keyboardHook;
        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="MsasSink"/> class.
        /// </summary>
        public MediaSink()
        {
            Trace.TraceInformation("MsasSink() Start");
            Trace.Indent();
            try
            {
                Trace.TraceInformation("Build version {0}", VersionInfo);
                m_portNumber = GetPortNumber(BASE_PORT);
                m_keyboardHook = new KeyboardHook();
                m_socketServer = new TcpSocketServer();
                m_eventSchedule = new EventSchedule();
            }
            catch (EventScheduleException ex)
            {
                //  This will happen if the user has not set up the EPG but it is not fatal
                Trace.TraceError(ex.ToString());
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
                throw;
            }
            finally
            {
                Trace.Unindent();
                Trace.TraceInformation("MsasSink() End");
            }
        }

        #region IMediaStatusSink Implementation
        /// <summary>
        /// Initialize Sink
        /// </summary>
		public void Initialize()
		{
            Trace.TraceInformation("Initialize() Start");
            Trace.Indent();
            try
            {
                //  Setup TCP socket server
                Trace.TraceInformation("Starting socket listener on port {0}", m_portNumber);
                m_socketServer.StartListening(m_portNumber);
                Trace.TraceInformation("Hooking socket connection events");
                m_socketServer.Connected += new EventHandler<SocketEventArgs>(SocketConnected);

                //  Hook into event schedule changes (if available)
                if (m_eventSchedule != null)
                {
                    Trace.TraceInformation("Hooking schedule change events");
                    m_eventSchedule.ScheduleEventStateChanged += new EventSchedule.EventStateChangedEventHandler(ScheduleEventStateChanged);
                }

                //  Hook into keyboard events
                Trace.TraceInformation("Hooking keyboard events");
                m_keyboardHook.KeyPress += new EventHandler<KeyboardHookEventArgs>(KeyPressEvent);

            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
                throw;
            }
            finally
            {
                Trace.Unindent();
                Trace.TraceInformation("Initialize() End");
            }
        }

        /// <summary>
        /// Creates a new media session.
        /// </summary>
        /// <returns>MsasSession</returns>
		public IMediaStatusSession CreateSession()
		{
            Trace.TraceInformation("CreateSession() Start");
            Trace.Indent();
            try
            {
                m_sessionCount++;
                Trace.TraceInformation("CreateSession() #{0}", m_sessionCount);
                return new MediaSession(m_sessionCount);
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
                throw;
            }
            finally
            {
                Trace.Unindent();
                Trace.TraceInformation("CreateSession() End");
            }
        }
        #endregion

        #region Event Handlers
        /// <summary>
        /// Handles the Connected event of the m_socketServer control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="VmcController.MceState.SocketEventArgs"/> instance containing the event data.</param>
        void SocketConnected(object sender, SocketEventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            Trace.TraceInformation("socketServer_Connected() Start");
            Trace.Indent();
            try
            {
                sb.AppendFormat(
                    "204 Connected (Build: {0} Clients: {1})\r\n",
                    VersionInfo, m_socketServer.Count);

                //  Provide current state information to the client
                if (!string.IsNullOrEmpty(MediaState.Volume))
                    sb.AppendFormat("Volume={0}\r\n", MediaState.Volume);
                if (!string.IsNullOrEmpty(MediaState.Mute))
                    sb.AppendFormat("Mute={0}\r\n", MediaState.Mute);
                if (MediaState.Page != MEDIASTATUSPROPERTYTAG.Unknown)
                    sb.AppendFormat("{0}=True\r\n", MediaState.Page);
                if (MediaState.MediaMode != MEDIASTATUSPROPERTYTAG.Unknown)
                    sb.AppendFormat("{0}=True\r\n", MediaState.MediaMode);
                if (MediaState.PlayRate != MEDIASTATUSPROPERTYTAG.Unknown)
                    sb.AppendFormat("{0}=True\r\n", MediaState.PlayRate);
                foreach (KeyValuePair<string, object> item in MediaState.MetaData)
                    sb.AppendFormat("{0}={1}\r\n", item.Key, item.Value);

                //  Send the data to the connected client
                Trace.TraceInformation(sb.ToString());
                m_socketServer.SendMessage(sb.ToString(), e.TcpClient);
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
            }
            finally
            {
                Trace.Unindent();
                Trace.TraceInformation("socketServer_Connected() End");
            }
        }

        /// <summary>
        /// Handles the KeyPress event of the m_keyboardHook control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="VmcController.MceState.KeyboardHookEventArgs"/> instance containing the event data.</param>
        void KeyPressEvent(object sender, KeyboardHookEventArgs e)
        {
            Trace.TraceInformation("keyboardHook_KeyPress() Start");
            Trace.Indent();
            try
            {
                //  Send the data to the clients
                Trace.TraceInformation("keyboardHook_KeyPress() event #{0}", e.vkCode);
                m_socketServer.SendMessage(String.Format(CultureInfo.InvariantCulture, "KeyPress={0}\r\n", e.vkCode));
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
            }
            finally
            {
                Trace.Unindent();
                Trace.TraceInformation("keyboardHook_KeyPress() End");
            }
        }

        /// <summary>
        /// Event schedule change event
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="Microsoft.MediaCenter.TV.Scheduling.ScheduleEventChangedEventArgs"/> instance containing the event data.</param>
        void ScheduleEventStateChanged(object sender, ScheduleEventChangedEventArgs e)
        {
            Trace.TraceInformation("ScheduleEventStateChanged() Start");
            Trace.Indent();
            try
            {
                foreach (ScheduleEventChange eventChange in e.Changes)
                {
                    Trace.TraceInformation("ScheduleEventStateChanged() event #{0}", eventChange.ScheduleEventId);
                    SocketServer.SendMessage(String.Format(CultureInfo.InvariantCulture, "Recording_{0}={1}\r\n", eventChange.NewState, eventChange.ScheduleEventId));
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.ToString());
            }
            finally
            {
                Trace.Unindent();
                Trace.TraceInformation("ScheduleEventStateChanged() End");
            }
        }
        #endregion

        #region Static Utility Methods
        /// <summary>
        /// Gets the TCP socket server.
        /// </summary>
        /// <value>The TCP socket server.</value>
        public static TcpSocketServer SocketServer
        {
            get { return m_socketServer; }
        }

        /// <summary>
        /// Gets the session count.
        /// </summary>
        /// <value>The session count.</value>
        public static int SessionCount
        {
            get { return MediaSink.m_sessionCount; }
        }

        /// <summary>
        /// Determine what TCP port to listen on
        /// </summary>
        /// <param name="basePort">The base port.</param>
        /// <returns>port number</returns>
        private static int GetPortNumber(int basePort)
        {
            string principalName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            int sessionId = Process.GetCurrentProcess().SessionId;
            Trace.TraceInformation("Windows Session #{0} Identity: {1}", sessionId, principalName);

            if (principalName.IndexOf("Mcx") > 0 && sessionId != 1)
                return basePort + int.Parse(principalName.Substring(principalName.LastIndexOf("Mcx") + 3), CultureInfo.InvariantCulture);
            else if (sessionId == 1)
                return basePort;
            else
                throw new InvalidOperationException("Unable to determine correct port number");
        }

        /// <summary>
        /// Gets the assembly version info.
        /// </summary>
        /// <value>The version info.</value>
        public static string VersionInfo
        {
            get {
                return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(3);
            }
        }
        #endregion
    }
}
