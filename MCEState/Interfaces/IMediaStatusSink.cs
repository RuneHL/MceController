/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 */
using System;
using System.Runtime.InteropServices;

namespace VmcController.MceState
{
	[ComImport, Guid("075FC453-F236-41DA-B90D-9FBB8BBDC101"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface IMediaStatusSink
	{
		[DispId(1)]
		void Initialize();
		[DispId(2)]
		IMediaStatusSession CreateSession();
	}
 }
