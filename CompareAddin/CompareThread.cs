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
		private string[] propertyList = 
			{
				"Business2TelephoneNumber",
				"FileAs","FullName","HomeAddress","HomeFaxNumber",
				"Email1Address","Email2Address","HomeTelephoneNumber",
				"Email1DisplayName","Email2DisplayName","JobTitle",
				"Email3Address","Email3DisplayName","MailingAddress",
				"BusinessFaxNumber","BusinessHomePage","MobileTelephoneNumber",
				"BusinessTelephoneNumber","CallbackTelephoneNumber","OtherAddress",
				"OtherFaxNumber","OtherTelephoneNumber","PrimaryTelephoneNumber",
				"CarTelephoneNumber","Department","Home2TelephoneNumber",
				"CompanyMainTelephoneNumber","FullNameAndCompany",
				"CompanyAndFullName","CompanyLastFirstNoSpace","CompanyLastFirstSpaceOnly",
				"BusinessAddress","CompanyName",
				"Account","Actions","Anniversary","AssistantName",
				"AssistantTelephoneNumber","Attachments",
				"AutoResolvedWinner","BillingInformation","Birthday","Body",
				"Categories","Children","Class","Companies",
				"ComputerNetworkName","Conflicts","ConversationIndex",
				"ConversationTopic","CustomerID",
				"DownloadState","Email1AddressType","Email2AddressType",
				"Email3AddressType","FormDescription","FTPSite",
				"Gender","GovernmentIDNumber",
				"HasPicture","Hobby","IMAddress","Importance",
				"Initials","InternetFreeBusyAddress",
				"ISDNNumber","Language","Links",
				"ManagerName","MarkForDownload","Mileage",
				"NetMeetingAlias","NetMeetingServer","NickName","NoAging",
				"OfficeLocation","OrganizationalIDNumber",
				"PagerNumber","Parent","PersonalHomePage",
				"Profession","RadioTelephoneNumber","ReferredBy",
				"SelectedMailingAddress","Sensitivity",
				"Session","Spouse","Subject",
				"Suffix","TelexNumber","Title",
				"TTYTDDTelephoneNumber","UnRead","UserCertificate","WebPage"
			};

		private Microsoft.Office.Interop.Outlook.Application application;

		public CompareThread(Microsoft.Office.Interop.Outlook.Application application)
		{
			this.application=application;
			//Thread thread = new Thread(new ThreadStart(Run));
			//thread.Start();
			Run();
		}

		private IDictionary badprop = new Hashtable();

		private void DoCompare(IList redundant, IList duplicateList, IDictionary cache, ItemPropertyHandler props, object obj)
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
						bool exact=false;
						UserProperties testprops = props.FetchUserProperties(obj);
						foreach (object known in list)
						{
							exact=true;
							UserProperties knownprops = props.FetchUserProperties(known);
							foreach (string prop in propertyList)
							{
								string knownprop = null;
								try
								{
									UserProperty check = knownprops.Find(prop,false);
									if (check!=null)
									{
										knownprop = (string)check.Value;
									}
								}
								catch {}
								string testprop = null;
								try
								{
									UserProperty check = testprops.Find(prop,false);
									if (check!=null)
									{
										testprop = (string)check.Value;
									}
								}
								catch {}
								if ((testprop==null)&&(knownprop==null))
								{
									continue;
								}
								if ((testprop!=null)&&(knownprop!=null))
								{
									testprop=testprop.Trim().ToLower();
									knownprop=knownprop.Trim().ToLower();
									if ((testprop==knownprop))
									{
										continue;
									}
								}
								if (badprop.Contains(prop))
								{
									badprop[prop]=((int)badprop[prop])+1;
								}
								else
								{
									badprop[prop]=1;
								}
								exact=false;
								break;
							}
							if (exact)
							{
								break;
							}
						}
						if (!exact)
						{
							list.Add(obj);
						}
						else
						{
							redundant.Add(obj);
						}
						if (list.Count==2)
						{
							duplicateList.Add(value);
						}
					}
					else
					{
						list = new ArrayList();
						list.Add(obj);
						cache.Add(value,list);
					}
				}
			}
		}

		private void Run()
		{
			CompareOptions options = new CompareOptions(application.GetNamespace("MAPI"));
			if (options.ShowDialog()==DialogResult.OK)
			{
				try
				{
					ItemPropertyHandler props = new ItemPropertyHandler(OlItemType.olContactItem,options.Field);

					IDictionary cache = new Hashtable();
					IList duplicateList = new ArrayList();
					IList redundant = new ArrayList();

					ProgressDialog progress = new ProgressDialog("Scanning Folders","Scanning folders. Please Wait.");
					progress.Value=0;
					progress.Maximum=options.Folder1.Items.Count;
					if (options.CompareMultipleFolders)
					{
						progress.Maximum+=options.Folder2.Items.Count;
					}
					progress.Show();
					foreach (object obj in options.Folder1.Items)
					{
						DoCompare(redundant,duplicateList,cache,props,obj);
						progress.Value++;
					}

					if (options.CompareMultipleFolders)
					{
						foreach (object obj in options.Folder2.Items)
						{
							DoCompare(redundant,duplicateList,cache,props,obj);
							progress.Value++;
						}
					}
					progress.Hide();
					progress.Close();
					progress.Dispose();

					if (MessageBox.Show("Found "+redundant.Count+" exact duplicates.\nWould you like to delete them?","Duplicates found",MessageBoxButtons.YesNo)==DialogResult.Yes)
					{
						progress = new ProgressDialog("Deleting Duplicates","Deleting duplicate contacts, please wait.");
						progress.Value=0;
						progress.Maximum=redundant.Count;
						progress.Show();

						foreach (ContactItem contact in redundant)
						{
							contact.Delete();
							progress.Value++;
						}

						progress.Hide();
						progress.Close();
						progress.Dispose();
					}

					/*IDictionaryEnumerator enumer = badprop.GetEnumerator();
					while(enumer.MoveNext())
					{
						MessageBox.Show(enumer.Key+" "+enumer.Value);
					}*/

					/*CompareResults results = new CompareResults(duplicateList,cache,props);
					results.Show();*/
				}
				catch (System.Exception e)
				{
					MessageBox.Show(e.Message);
				}
			}
		}
	}
}
