/*
 * $LastChangedBy$
 * $HeadURL$
 * $Date$
 * $Revision: 6 $
 */

using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Collections;

namespace CompareAddin
{
	public class OutlookCompare
	{
		private Application application;
		private IList objectCache = new ArrayList();

		public OutlookCompare(Connect connect)
		{
			connect.OutlookStarted+=new OutlookStartedHandler(connect_OutlookStarted);
			connect.OutlookShuttingDown+=new OutlookShuttingDownHandler(connect_OutlookShuttingDown);
		}

		private void connect_OutlookStarted(Application application)
		{
			this.application=application;
			foreach (Explorer explorer in application.Explorers)
			{
				CreateMenu(explorer);
			}
		}

		private void connect_OutlookShuttingDown(Application application)
		{
			this.application=null;
		}

		private void CreateMenu(Explorer explorer)
		{
			int pos = explorer.CommandBars["Menu Bar"].Controls["Help"].Index;
			CommandBarPopup menu = (CommandBarPopup)explorer.CommandBars["Menu Bar"].Controls.Add(MsoControlType.msoControlPopup,System.Reflection.Missing.Value,System.Reflection.Missing.Value,pos,true);
			menu.OnAction = "!<OutlookCompare.Connect>";
			menu.Caption = "OutlookCompare";
			CommandBarButton comparemenu = (CommandBarButton)menu.Controls.Add(MsoControlType.msoControlButton,System.Reflection.Missing.Value,System.Reflection.Missing.Value,System.Reflection.Missing.Value,true);
			comparemenu.OnAction="!<MaillistManager.Connect>";
			comparemenu.Caption="Compare Folders...";
			comparemenu.Click+=new _CommandBarButtonEvents_ClickEventHandler(comparemenu_Click);

			objectCache.Add(comparemenu);
		}

		private void comparemenu_Click(CommandBarButton Ctrl, ref bool CancelDefault)
		{
			new CompareThread(application);
		}
	}
}
