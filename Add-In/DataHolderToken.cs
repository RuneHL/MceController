﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMPLib;

namespace VmcController.AddIn
{
    /// <summary>
    /// This class serves to store data about the current SocketAsyncEventArgs object in the UserToken property
    /// </summary>
    class DataHolderToken
    {
        //Variables for handling the network request thread
        internal NowPlayingList nowPlaying;
        internal MediaItem currentMedia;        
        internal OpResult opResult;

        internal Byte[] dataMessageReceived;
        internal Byte[] dataToSend;

        internal string httpRequest;

        //The offset in SocketAsyncEventArgs.Offset
        internal int receiveBufferOffset;

        //This is the receive buffer size + the offset above
        internal int sendBufferOffset;

        internal int messageBytesReceived = 0;

        internal int sendBytesRemainingCount;
        internal int bytesSentAlreadyCount;


        public DataHolderToken(int rOffset, int receiveBufferSize)
        {
            receiveBufferOffset = rOffset;
            sendBufferOffset = rOffset + receiveBufferSize;
        }

        public void Reset()
        {
            nowPlaying = null;
            currentMedia = null;
            opResult = null;

            httpRequest = null;
            dataMessageReceived = null;
            dataToSend = null;
            messageBytesReceived = 0;
        }
    }
}
