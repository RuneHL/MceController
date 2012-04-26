/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 */
using System;
using System.Runtime.InteropServices;

namespace VmcController.MceState
{
	[ComImport, Guid("A70D81F2-C9D2-4053-AF0E-CDEA39BDD1AD"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface IMediaStatusSession
	{
		[DispId(1)]
		void MediaStatusChange(MEDIASTATUSPROPERTYTAG[] tags, object[] properties);
		[DispId(2)]
		void Close();
	}
}
