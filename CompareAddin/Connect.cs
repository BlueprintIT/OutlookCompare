/*
 * $LastChangedBy$
 * $HeadURL$
 * $Date$
 * $Revision$
 */

using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Extensibility;
using System.Runtime.InteropServices;
	
namespace CompareAddin
{
	#region Read me for Add-in installation and setup information.
	// When run, the Add-in wizard prepared the registry for the Add-in.
	// At a later time, if the Add-in becomes unavailable for reasons such as:
	//   1) You moved this project to a computer other than which is was originally created on.
	//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
	//   3) Registry corruption.
	// you will need to re-register the Add-in by building the MyAddin21Setup project 
	// by right clicking the project in the Solution Explorer, then choosing install.
	#endregion
	
	public delegate void OutlookStartedHandler(Application application);

	public delegate void OutlookShuttingDownHandler(Application application);

	[GuidAttribute("A0C190C0-C85A-4172-8543-DFFE9E9AF87E"), ProgId("CompareAddin.Connect")]
	public class Connect : Extensibility.IDTExtensibility2
	{
		private Application applicationObject;
		private object addInInstance;

		public event OutlookStartedHandler OutlookStarted;
		public event OutlookShuttingDownHandler OutlookShuttingDown;

		public Connect()
		{
			new OutlookCompare(this);
		}

		public Application OutlookApplication
		{
			get
			{
				return applicationObject;
			}
		}

		public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
		{
			applicationObject = (Application)application;
			addInInstance = addInInst;

			if(connectMode != Extensibility.ext_ConnectMode.ext_cm_Startup)
			{
				OnStartupComplete(ref custom);
			}
		}

		public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
		{
			if(disconnectMode != Extensibility.ext_DisconnectMode.ext_dm_HostShutdown)
			{
				OnBeginShutdown(ref custom);
			}
		}

		public void OnAddInsUpdate(ref System.Array custom)
		{
		}

		public void OnStartupComplete(ref System.Array custom)
		{
			if (OutlookStarted!=null)
			{
				OutlookStarted(applicationObject);
			}
		}

		public void OnBeginShutdown(ref System.Array custom)
		{
			if (OutlookShuttingDown!=null)
			{
				OutlookShuttingDown(applicationObject);
			}
			OutlookStarted=null;
			OutlookShuttingDown=null;
			applicationObject=null;
			addInInstance=null;
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}
	}
}