/*
 * $LastChangedBy$
 * $HeadURL$
 * $Date$
 * $Revision$
 */

using System;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.Collections;

namespace CompareAddin
{
	public class CompareThread
	{
		private Microsoft.Office.Interop.Outlook.Application application;

		public CompareThread(Microsoft.Office.Interop.Outlook.Application application)
		{
			this.application=application;
			//Thread thread = new Thread(new ThreadStart(Run));
			//thread.Start();
			Run();
		}

		private void DoCompare(IList duplicateList, IDictionary cache, ItemPropertyHandler props, object obj)
		{
			if (props.IsCorrectType(obj))
			{
				string value = props.FetchIndexProperty(obj);
				if ((value!=null)&&(value.Length>0))
				{
					IList list;
					if (cache.Contains(value))
					{
						list = (IList)cache[value];
					}
					else
					{
						list = new ArrayList();
						cache.Add(value,list);
					}
					list.Add(obj);
					if (list.Count==2)
					{
						duplicateList.Add(value);
					}
				}
			}
		}

		private void Run()
		{
			CompareOptions options = new CompareOptions(application.GetNamespace("MAPI"));
			if (options.ShowDialog()==DialogResult.OK)
			{
				ItemPropertyHandler props = new ItemPropertyHandler(OlItemType.olContactItem,"Email1Address");

				IDictionary cache = new Hashtable();
				IList duplicateList = new ArrayList();

				ProgressDialog progress = new ProgressDialog("Scanning Folders","Scanning folders. Please Wait.");
				progress.Value=0;
				progress.Maximum=options.Folder1.Items.Count;
				progress.Show();
				foreach (object obj in options.Folder1.Items)
				{
					DoCompare(duplicateList,cache,props,obj);
					progress.Value++;
				}

				if (options.CompareMultipleFolders)
				{
					progress.Value=0;
					progress.Maximum=options.Folder2.Items.Count;
					foreach (object obj in options.Folder2.Items)
					{
						DoCompare(duplicateList,cache,props,obj);
						progress.Value++;
					}
				}
				progress.Hide();
				progress.Close();
				progress.Dispose();

				CompareResults results = new CompareResults(duplicateList,cache,props);
				results.Show();
			}
		}
	}
}
