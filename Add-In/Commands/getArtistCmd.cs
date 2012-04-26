/*
 * This module is build on top of on J.Bradshaw's vmcController
 * Implements audio media library functions
 * 
 * Copyright (c) 2008 Anthony Jones
 * 
 * Portions copyright (c) 2007 Jonathan Bradshaw
 * 
 * This software code module is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely.
 * 
 * History:
 * 2008-12-18 Created by Anthony Jones
 * 
 */

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Microsoft.MediaCenter;

namespace VmcController.AddIn.Commands
{
    /// <summary>
    /// Summary description for getArtists commands.
    /// </summary>
    public class artistCmd : ICommand
    {
        private WMPLib.WindowsMediaPlayer Player = null;
        private WMPLib.IWMPMedia media;
        private WMPLib.IWMPPlaylist mediaPlaylist;

        private bool play = false;
        private bool queue = false;
        private bool debug = false;

        public static int show_artists = 1;
        public static int show_albums = 2;
        public static int show_songs = 4;
        public static int show_genre = 8;

        private int showWhat = show_artists;
        private int which_command = -1;

        public static int by_all = 0;
        public static int by_track = 1;
        public static int by_artist = 2;
        public static int by_album = 3;
        private int showBy = by_all;

        private Dictionary<string, string> m_templates = new Dictionary<string, string>();

        public artistCmd(int i)
        {
            which_command = i;
        }

        public artistCmd(bool bplay)
        {
            play = bplay;
            loadTemplate();
        }

        public artistCmd(bool bplay, bool bqueue)
        {
            play = bplay;
            queue = bqueue;
            loadTemplate();
        }

        public artistCmd(bool bplay, bool bqueue, bool bdebug)
        {
            play = bplay;
            queue = bqueue;
            this.debug = bdebug;
            loadTemplate();
        }

        public artistCmd(bool bplay, bool bqueue, bool bdebug, int iShowWhat)
        {
            play = bplay;
            queue = bqueue;
            this.debug = bdebug;
            this.showWhat = iShowWhat;
            loadTemplate();
        }
        public artistCmd(bool bplay, bool bqueue, bool bdebug, int iShowWhat, int iShowBy)
        {
            play = bplay;
            queue = bqueue;
            this.debug = bdebug;
            this.showWhat = iShowWhat;
            this.showBy = iShowBy;
            loadTemplate();
        }

        private bool loadTemplate()
        {
            bool ret = true;

            try
            {
                Regex re = new Regex("(?<lable>.+?)\t+(?<format>.*$?)");
                StreamReader fTemplate = File.OpenText("artist.template");
                string sIn = null;
                while ((sIn = fTemplate.ReadLine()) != null)
                {
                    Match match = re.Match(sIn);
                    if (match.Success) m_templates.Add(match.Groups["lable"].Value, match.Groups["format"].Value);
                }
                fTemplate.Close();
            }
            catch (Exception) { ret = false; }

            return ret;
        }

        private struct state
        {
            public string artist;
            public string nextArtist;
            public int artistIndex;

            public int artistTrackCount;
            public int artistAlbumCount;
            public string artistAlbumList;
            public int artistCount;

            public string album;
            public string nextAlbum;
            public int albumIndex;
            public int albumTrackCount;
            public string albumYear;
            public string albumGenre;
            public string albumImage;

            public string song;
            public string songTrackNumber;
            public string songLength;
            public string songLocation;

            public int trackCount;
            public int albumCount;

            public void init()
            {
                artist = "";
                nextArtist = "";
                artistIndex = 0;
                artistTrackCount = 0;
                artistAlbumCount = 0;
                artistAlbumList = "";
                artistCount = 0;

                album = "";
                nextAlbum = "";
                albumIndex = -1;
                albumTrackCount = 0;
                albumYear = "";
                albumGenre = "";
                albumImage = "";

                song = "";
                songTrackNumber = "";
                songLength = "";
                songLocation = "";

                trackCount = 0;
                albumCount = 0;
            }

            public void resetArtist()
            {
                artist = "";
                artistTrackCount = 0;
                artistAlbumCount = 0;
                artistAlbumList = "";
            }

            public void resetAlbum()
            {
                album = "";
                albumIndex = -1;
                albumTrackCount = 0;
                songTrackNumber = "";
                songLength = "";
                albumYear = "";
                albumGenre = "";
                albumImage = "";
            }

            public void albumList_add(string s)
            {
                if (artistAlbumList.Length == 0) artistAlbumList = s;
                else artistAlbumList = artistAlbumList + ", " + s;
            }

            public void albumGenre_add(string s)
            {
                if (s == "") return;
                if (albumGenre.IndexOf(s) >= 0) return;
                if (albumGenre.Length == 0) albumGenre = s;
                else albumGenre = albumGenre + ", " + s;
            }

            public void findAlbumCover(string url)
            {
                string s = "";
                if (albumImage.Length > 0) return;

                try
                {
                    s = Path.GetDirectoryName(url) + @"\Folder.jpg";
                    if (File.Exists(s)) albumImage = s;
                    else
                    {
                        s = Path.GetDirectoryName(url) + @"\AlbumArtSmall.jpg";
                        if (File.Exists(s)) albumImage = s;
                    }
                }
                finally { }
            }

            public string replacer(string s)
            {
                s = s.Replace("%artist%", artist);
                s = s.Replace("%nextArtist%", nextArtist);
                s = s.Replace("%artistIndex%", String.Format("{0}", artistIndex));
                s = s.Replace("%artistTrackCount%", String.Format("{0}", artistTrackCount));
                s = s.Replace("%artistAlbumCount%", String.Format("{0}", artistAlbumCount));
                s = s.Replace("%artistCount%", String.Format("{0}", artistCount));
                s = s.Replace("%artistAlbumList%", artistAlbumList);

                s = s.Replace("%album%", album);
                s = s.Replace("%nextAlbum%", nextAlbum);
                s = s.Replace("%albumIndex%", String.Format("{0}", albumIndex));
                s = s.Replace("%albumTrackCount%", String.Format("{0}", albumTrackCount));
                s = s.Replace("%albumYear%", albumYear);
                s = s.Replace("%albumGenre%", albumGenre);
                if (albumImage.Length > 0) s = s.Replace("%albumImage%", albumImage);
                else s = s.Replace("%albumImage%", "NoAlbumImage");

                s = s.Replace("%song%", song);
                s = s.Replace("%songPath%", songLocation);
                s = s.Replace("%songTrackNumber%", songTrackNumber);
                s = s.Replace("%songLength%", songLength);

                s = s.Replace("%trackCount%", String.Format("{0}", trackCount));
                s = s.Replace("%albumCount%", String.Format("{0}", albumCount));

                return s;
            }
        }
        private state the_state;

        private bool templateOut(string template, OpResult or, int idx, double elapsed)
        {
            string tmp = "";

            if (!m_templates.ContainsKey(template)) return false;

            tmp = m_templates[template];
            tmp = tmp.Replace("%idx%", String.Format("{0}", idx));
            tmp = tmp.Replace("%elapsed_time%", String.Format("{0}", elapsed));
            tmp = tmp.Replace("\\r", "\r");
            tmp = tmp.Replace("\\n", "\n");
            tmp = tmp.Replace("\\t", "\t");
            tmp = tmp.Replace("{", "{{");
            tmp = tmp.Replace("}", "}}");
            tmp = the_state.replacer(tmp);
            if (tmp.TrimStart(' ').Length > 0) or.AppendFormat(tmp);

            return true;
        }

        private static Regex custom_regex = new Regex("-(?<index>\\d\\d\\d+)\\s*(?<remainder>.*)");
        private static Regex index_regex = new Regex("<(?<index>\\d+)>");

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            string s = "";
            s = "[-help] [-custom#] [artist_name_filter | <index#>] - list / play from audio collection";
            return s;
        }

        public OpResult showHelp(OpResult or)
        {
            or.AppendFormat("list-artists [artist_name_filter] - lists all matching artists (Note very slow!)");
            or.AppendFormat("list-artist-songs <index#> - lists all songs for the specified artist");
            or.AppendFormat("list-artist-albums <index#> - lists all albums for the specified artist");
            or.AppendFormat("list-albums - lists all albums");
            or.AppendFormat("list-album-songs <index#> - lists all songs for the specified album");
            or.AppendFormat("list-all-custom [-custom#] - Uses the specified custom format against all audio items");
            or.AppendFormat("list-artist-custom [-custom#] <index#> - Uses the specified custom format for the specified artist");
            or.AppendFormat("list-album-custom [-custom#] <index#> - Uses the specified custom format for the specified album");
            or.AppendFormat("list-song-custom [-custom#] <index#> - Uses the specified custom format for the specified song");
            or.AppendFormat("play-audio-artist <index#> - plays all songs for the specified artist");
            or.AppendFormat("play-audio-album <index#> - plays all songs for the specified album");
            or.AppendFormat("play-audio-song <index#> - plays the specified song");
            or.AppendFormat("queueaudio-artist <index#> - Adds all songs for the specified artist to the queue");
            or.AppendFormat("queueaudio-album <index#> - Adds all songs for the specified album to the queue");
            or.AppendFormat("queueaudio-song <index#> - Adds the specified song to the queue");
            or.AppendFormat(" ");
            or.AppendFormat("Where:");
            or.AppendFormat("     artist_name_filter: a filter string (\"ab\" would only return artists starting with \"Ab\")");
            or.AppendFormat("          NOTE: Using an artist_name_filter is very slow! May take minutes to return with 10K tracks!!");
            or.AppendFormat("     custom#: a number > 100 which specifies a custom format in the artist.template file");
            or.AppendFormat("     index#: a track index (from the default \"list-xxx\" commands)");
            or.AppendFormat("          NOTE: Using the track index is very fast! Use the list-xxx commands to retrieve the indexes");
            or.AppendFormat(" ");
            or.AppendFormat("Parameter Notes:");
            or.AppendFormat("     The index number is just an index into the complete collection of songs and may change - it is not static!");
            or.AppendFormat("     Almost of the parameters can be used with any command - the above just shows the parameters that make sense.");
            or.AppendFormat(" ");
            or.AppendFormat("Custom Formatting Notes:");
            or.AppendFormat("     All custom formats must be defined in the \"artist.template\" file");
            or.AppendFormat("     The \"artist.template\" file must be manually coppied to the ehome directory (usually C:\\Windows\\ehome)");
            or.AppendFormat("     The \"artist.template\" file contains notes / examples on formatting");
            or.AppendFormat("     The built in formats may also be overwritten via the \"artist.template\" file");

            return or;
        }

        public OpResult listArtistsOnly(OpResult or, string filter)
        {
            WMPLib.IWMPMediaCollection2 collection = null;

            collection = (WMPLib.IWMPMediaCollection2)Player.mediaCollection;
            WMPLib.IWMPQuery query = collection.createQuery();
            WMPLib.IWMPStringCollection artists = null;

            if (filter.Length > 0) query.addCondition("Artist", "Contains", filter);
            artists = collection.getStringCollectionByQuery("Artist", query, "Audio", "", true);
            for (int j = 0; j < artists.count; j++)
            {
                or.AppendFormat("artist={0}", artists.Item(j));
            }

            return or;
        }

        public WMPLib.IWMPPlaylist byArtist(string filter)
        {
            WMPLib.IWMPMediaCollection2 collection = null;

            collection = (WMPLib.IWMPMediaCollection2)Player.mediaCollection;
            WMPLib.IWMPQuery query = collection.createQuery();

            if (filter.Length > 0)
            {
                query.addCondition("Artist", "BeginsWith", filter);
                query.beginNextGroup();
                query.addCondition("Artist", "Contains", " " + filter);
            }
            return collection.getPlaylistByQuery(query, "Audio", "", false);
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
            bool bFirst = !queue;
            bool bByIndex = false;
            int idx = 0;

            opResult.StatusCode = OpStatusCode.Ok;
            try
            {
                if (Player == null) Player = new WMPLib.WindowsMediaPlayer();

                if (param.IndexOf("-help") >= 0)
                {
                    opResult = showHelp(opResult);
                    return opResult;
                }

                DateTime startTime = DateTime.Now;

                // Use custom format?
                Match custom_match = custom_regex.Match(param);
                if (custom_match.Success)
                {
                    showWhat = System.Convert.ToInt32(custom_match.Groups["index"].Value, 10);
                    param = custom_match.Groups["remainder"].Value;
                }

                Match match = index_regex.Match(param);
                if (match.Success)
                {
                    idx = System.Convert.ToInt32(match.Groups["index"].Value, 10);
                    bByIndex = true;
                    param = "";
                }
                else
                {
                    idx = 0;
                    param = param.ToLower();
                }

                if (showWhat == 0)
                {
                    opResult = listArtistsOnly(opResult, param);
                    TimeSpan duration = DateTime.Now - startTime;
                    opResult.AppendFormat("elapsed_time={0}", duration.TotalSeconds);
                    return opResult;
                }
                if (param.Length > 0) mediaPlaylist = byArtist(param);
                else mediaPlaylist = Player.mediaCollection.getByAttribute("MediaType", "Audio");

                bool bArtistMatch = false;
                the_state.init();

                int iCount = mediaPlaylist.count;

                string keyArtist = "";
                string keyAlbum = "";

                // Header:
                templateOut(String.Format("{0}H", showWhat), opResult, 0, -1);

                for (int x = idx; x < iCount; x++)
                {
                    media = mediaPlaylist.get_Item(x);
                    the_state.nextArtist = media.getItemInfo("WM/AlbumArtist");
                    if (the_state.nextArtist == "") the_state.nextArtist = media.getItemInfo("Author");
                    the_state.nextAlbum = media.getItemInfo("WM/AlbumTitle");
                    if (bByIndex && x == idx)
                    {
                        keyArtist = the_state.nextArtist;
                        keyAlbum = the_state.nextAlbum;
                    }
                    else
                    {
                        if (showBy == by_track && keyArtist.Length > 0 && x != idx) break;
                        if (showBy == by_artist && keyArtist.Length > 0 && the_state.nextArtist != keyArtist) break;
                        if (showBy == by_album && keyArtist.Length > 0 && (the_state.nextAlbum != keyAlbum || the_state.nextArtist != keyArtist)) break;
                    }

                    if (the_state.nextArtist != the_state.artist)   // New artist?
                    {
                        if (bArtistMatch) // Did last artist match?
                        {
                            // Close album:
                            if (the_state.album.Length > 0)
                            {
                                if (!templateOut(String.Format("{0}.{1}-", showWhat, show_albums), opResult, x, -1) && ((showWhat & show_albums) != 0) && the_state.albumTrackCount > 0)
                                    opResult.AppendFormat("    Album_song_count: {0}", the_state.albumTrackCount);
                            }
                            the_state.resetAlbum();
                            // Close artist"
                            if (!templateOut(String.Format("{0}.{1}-", showWhat, show_artists), opResult, x, -1))
                            {
                                if (the_state.artistAlbumCount > 0) opResult.AppendFormat("  Artist_album_count: {0}", the_state.artistAlbumCount);
                                if (the_state.artistTrackCount > 0) opResult.AppendFormat("  Artist_song_count: {0}", the_state.artistTrackCount);
                            }
                            the_state.resetArtist();
                        }
                        if (the_state.nextArtist.ToLower().StartsWith(param))
                        {
                            bArtistMatch = true;
                            the_state.artistCount += 1;

                            the_state.artistIndex = x;

                            if (keyArtist.Length <= 0)
                            {
                                keyArtist = the_state.nextArtist;
                                keyAlbum = the_state.nextAlbum;
                            }
                        }
                        else bArtistMatch = false;
                        the_state.artist = the_state.nextArtist;
                        if (bArtistMatch)
                        {
                            // Open Artist
                            if (!templateOut(String.Format("{0}.{1}+", showWhat, show_artists), opResult, x, -1) && ((showWhat & show_artists) != 0))
                                opResult.AppendFormat("Artist:<{0}> \"{1}\"", x, the_state.artist);
                        }
                    }
                    if (bArtistMatch)
                    {
                        the_state.artistTrackCount += 1;
                        the_state.trackCount += 1;
                        if (the_state.nextAlbum != the_state.album)
                        {
                            if (the_state.album.Length > 0)
                            {
                                // Close album
                                if (!templateOut(String.Format("{0}.{1}-", showWhat, show_albums), opResult, x, -1) && ((showWhat & show_albums) != 0) && the_state.albumTrackCount > 0)
                                    opResult.AppendFormat("    Album_song_count: {0}", the_state.albumTrackCount);
                            }
                            the_state.resetAlbum();
                            the_state.album = the_state.nextAlbum;
                            if (the_state.nextAlbum.Length > 0)
                            {
                                the_state.albumIndex = x;
                                the_state.albumCount += 1;
                                the_state.artistAlbumCount += 1;
                                the_state.albumList_add(the_state.nextAlbum);
                                the_state.albumYear = media.getItemInfo("WM/Year");
                                the_state.albumGenre = media.getItemInfo("WM/Genre");
                                the_state.findAlbumCover(media.sourceURL);
                                // Open album
                                if (!templateOut(String.Format("{0}.{1}+", showWhat, show_albums), opResult, x, -1) && ((showWhat & show_albums) != 0))
                                    opResult.AppendFormat("  Album:<{0}> \"{1}\"", x, the_state.nextAlbum);
                            }
                        }
                        the_state.album = the_state.nextAlbum;
                        the_state.albumTrackCount += 1;
                        the_state.song = media.getItemInfo("Title");
                        the_state.songLocation = media.sourceURL;
                        the_state.songTrackNumber = media.getItemInfo("WM/TrackNumber");
                        the_state.songLength = media.durationString;
                        if (the_state.albumYear == "" || the_state.albumYear.Length < 4) the_state.albumYear = media.getItemInfo("WM/OriginalReleaseYear");
                        the_state.albumGenre_add(media.getItemInfo("WM/Genre"));
                        the_state.findAlbumCover(the_state.songLocation);

                        // Open / close song
                        if (!templateOut(String.Format("{0}.{1}-", showWhat, show_songs), opResult, x, -1) && ((showWhat & show_songs) != 0) && the_state.song.Length > 0)
                            opResult.AppendFormat("    Song:<{0}> \"{1}\" ({2})", x, the_state.song, the_state.songLength);
                        templateOut(String.Format("{0}.{1}+", showWhat, show_songs), opResult, x, -1);

                        // Play / queue song?
                        if (play)
                        {
                            PlayMediaCmd pmc;
                            pmc = new PlayMediaCmd(MediaType.Audio, !bFirst);
                            bFirst = false;
                            opResult = pmc.Execute(media.sourceURL);
                        }
                    }
                }
                if (play) opResult.AppendFormat("Queued {0} songs to play", the_state.trackCount);
                else
                {
                    // Close album:
                    if (the_state.album.Length > 0)
                    {
                        if (!templateOut(String.Format("{0}.{1}-", showWhat, show_albums), opResult, the_state.trackCount, -1) && ((showWhat & show_albums) != 0) && the_state.albumTrackCount > 0)
                            opResult.AppendFormat("    Album_song_count: {0}", the_state.albumTrackCount);
                    }
                    the_state.resetAlbum();
                    // Close artist:
                    if (!templateOut(String.Format("{0}.{1}-", showWhat, show_artists), opResult, the_state.trackCount, -1))
                    {
                        if (the_state.artistAlbumCount > 0) opResult.AppendFormat("  Artist_album_count: {0}", the_state.artistAlbumCount);
                        if (the_state.artistTrackCount > 0) opResult.AppendFormat("  Artist_song_count: {0}", the_state.artistTrackCount);
                    }
                    the_state.resetArtist();
                }
                // Footer:
                TimeSpan elapsed_duration = DateTime.Now - startTime;
                if (!templateOut(String.Format("{0}F", showWhat), opResult, the_state.trackCount, elapsed_duration.TotalSeconds))
                {
                    if (the_state.artistCount > 0) opResult.AppendFormat("Artist_count: {0}", the_state.artistCount);
                    if (the_state.albumCount > 0) opResult.AppendFormat("Album_count: {0}", the_state.albumCount);
                    if (the_state.trackCount > 0) opResult.AppendFormat("Song_count: {0}", the_state.trackCount);
                    opResult.AppendFormat("elapsed_time={0}", elapsed_duration.TotalSeconds);
                }

            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = ex.Message;
            }
            return opResult;
        }

        private String out_count(int map, String label, int count, String key)
        {
            if ((showWhat & map) == map && key.Length > 0 && count > 0)
            {
                return (String.Format("{0}:{1}", label, count));
            }
            else return "";
        }

        #endregion
    }
}
