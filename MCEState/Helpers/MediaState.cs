/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;

namespace VmcController.MceState
{
    /// <summary>
    /// Tracks state information
    /// </summary>
    public static class MediaState
    {
        private static MEDIASTATUSPROPERTYTAG m_mediaMode = MEDIASTATUSPROPERTYTAG.Unknown;
        private static MEDIASTATUSPROPERTYTAG m_playRate = MEDIASTATUSPROPERTYTAG.Unknown;
        private static MEDIASTATUSPROPERTYTAG m_page = MEDIASTATUSPROPERTYTAG.Unknown;
        private static Dictionary<string, object> m_metaData = new Dictionary<string, object>();
        private static string m_volume = string.Empty;
        private static string m_mute = string.Empty;

        /// <summary>
        /// Gets the volume.
        /// </summary>
        /// <value>The volume.</value>
        public static string Volume
        {
            get { return MediaState.m_volume; }
        }

        /// <summary>
        /// Gets the mute.
        /// </summary>
        /// <value>The mute.</value>
        public static string Mute
        {
            get { return MediaState.m_mute; }
        }

        /// <summary>
        /// Gets the meta data.
        /// </summary>
        /// <value>The meta data.</value>
        public static Dictionary<string, object> MetaData
        {
            get { return MediaState.m_metaData; }
        }

        /// <summary>
        /// Gets the current page.
        /// </summary>
        /// <value>The page.</value>
        public static MEDIASTATUSPROPERTYTAG Page
        {
            get { return MediaState.m_page; }
        }

        /// <summary>
        /// Gets the current play rate.
        /// </summary>
        /// <value>The play rate.</value>
        public static MEDIASTATUSPROPERTYTAG PlayRate
        {
            get { return MediaState.m_playRate; }
        }

        /// <summary>
        /// Gets the current media mode.
        /// </summary>
        /// <value>The media mode.</value>
        public static MEDIASTATUSPROPERTYTAG MediaMode
        {
            get { return MediaState.m_mediaMode; }
        }

        /// <summary>
        /// Updates the state.
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="property">The property.</param>
        public static void UpdateState(MEDIASTATUSPROPERTYTAG tag, object property)
        {
            if (property == null) return;

            switch (tag)
            {
                //  Volume
                case MEDIASTATUSPROPERTYTAG.Volume:
                    m_volume = property.ToString();
                    break;

                //  Mute
                case MEDIASTATUSPROPERTYTAG.Mute:
                    m_mute = property.ToString();
                    break;

                //  Current Navigation Page
                case MEDIASTATUSPROPERTYTAG.FS_DVD:
                case MEDIASTATUSPROPERTYTAG.FS_Guide:
                case MEDIASTATUSPROPERTYTAG.FS_Home:
                case MEDIASTATUSPROPERTYTAG.FS_Music:
                case MEDIASTATUSPROPERTYTAG.FS_Photos:
                case MEDIASTATUSPROPERTYTAG.FS_Radio:
                case MEDIASTATUSPROPERTYTAG.FS_RecordedShows:
                case MEDIASTATUSPROPERTYTAG.FS_TV:
                case MEDIASTATUSPROPERTYTAG.FS_Unknown:
                case MEDIASTATUSPROPERTYTAG.FS_Videos:
                    if ((bool)property == true) m_page = tag;
                    break;

                //  Play Rates
                case MEDIASTATUSPROPERTYTAG.Play:
                case MEDIASTATUSPROPERTYTAG.Stop:
                case MEDIASTATUSPROPERTYTAG.Pause:
                case MEDIASTATUSPROPERTYTAG.FF1:
                case MEDIASTATUSPROPERTYTAG.FF2:
                case MEDIASTATUSPROPERTYTAG.FF3:
                case MEDIASTATUSPROPERTYTAG.Rewind1:
                case MEDIASTATUSPROPERTYTAG.Rewind2:
                case MEDIASTATUSPROPERTYTAG.Rewind3:
                case MEDIASTATUSPROPERTYTAG.SlowMotion1:
                case MEDIASTATUSPROPERTYTAG.SlowMotion2:
                case MEDIASTATUSPROPERTYTAG.SlowMotion3:
                    if ((bool)property == true)
                        m_playRate = tag;
                    break;

                //  Current Media Mode
                case MEDIASTATUSPROPERTYTAG.StreamingContentAudio:
                case MEDIASTATUSPROPERTYTAG.StreamingContentVideo:
                case MEDIASTATUSPROPERTYTAG.PVR:
                case MEDIASTATUSPROPERTYTAG.TVTuner:
                case MEDIASTATUSPROPERTYTAG.CD:
                case MEDIASTATUSPROPERTYTAG.DVD:
                    if ((bool)property == true)
                    {
                        m_metaData.Clear();
                        m_mediaMode = tag;
                    }
                    break;

                default:
                    m_metaData[tag.ToString()] = property;
                    break;
            }
        }
    }
}
