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

		private UserProperties FetchUserProperties(object obj, OlItemType type)
		{
			if (type==OlItemType.olContactItem)
			{
				if (obj is ContactItem)
				{
					return ((ContactItem)obj).UserProperties;
				}
				else
				{
					return null;
				}
			}
			else
			{
				return null;
			}
		}

		private string Normalise(string value)
		{
			return value.Trim().ToLower();
		}

		private void DoCompare(IList duplicateList, IDictionary cache, object obj, OlItemType type, string property)
		{
			UserProperties props = FetchUserProperties(obj,type);
			if (props!=null)
			{
				UserProperty prop = props.Find(property,false);
				if (prop!=null)
				{
					string value = Normalise((string)prop.Value);
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
				IDictionary cache = new Hashtable();
				IList duplicateList = new ArrayList();

				ProgressDialog progress = new ProgressDialog("Scanning Folders","Scanning folders. Please Wait.");
				progress.Value=0;
				progress.Maximum=options.Folder1.Items.Count;
				progress.Show();
				foreach (object obj in options.Folder1.Items)
				{
					DoCompare(duplicateList,cache,obj,OlItemType.olContactItem,"Email1Address");
					progress.Value++;
				}

				if (options.CompareMultipleFolders)
				{
					progress.Value=0;
					progress.Maximum=options.Folder2.Items.Count;
					foreach (object obj in options.Folder2.Items)
					{
						DoCompare(duplicateList,cache,obj,OlItemType.olContactItem,"Email1Address");
						progress.Value++;
					}
				}
				progress.Hide();
				progress.Close();
				progress.Dispose();

				MessageBox.Show("Non-duplicates: "+(cache.Count-duplicateList.Count)+" Duplicates: "+duplicateList.Count);
			}
		}
	}
}
