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
				"Account","Actions","Anniversary",
				"Application","AssistantName",
				"AssistantTelephoneNumber","Attachments",
"AutoResolvedWinner","BillingInformation","Birthday",
"Body","Business2TelephoneNumber","BusinessAddress",
"BusinessAddressCity","BusinessAddressCountry",
				"BusinessAddressPostalCode","BusinessAddressPostOfficeBox",
"BusinessAddressState","BusinessAddressStreet",
"BusinessFaxNumber","BusinessHomePage",
"BusinessTelephoneNumber","CallbackTelephoneNumber",
"CarTelephoneNumber",
				"Categories","Children","Class",
"Companies","CompanyAndFullName",
"CompanyLastFirstNoSpace","CompanyLastFirstSpaceOnly",
"CompanyMainTelephoneNumber","CompanyName",
"ComputerNetworkName","Conflicts","ConversationIndex",
"ConversationTopic","CreationTime","CustomerID",
"Department","DownloadState","Email1Address",
"Email1AddressType","Email1DisplayName","Email1EntryID",
"Email2Address",
				"Email2AddressType","Email2DisplayName",
"Email2EntryID","Email3Address","Email3AddressType",
"Email3DisplayName","Email3EntryID","EntryID",
"FileAs","FirstName","FormDescription",
"FTPSite","FullName","FullNameAndCompany",
"Gender","GetInspector","GovernmentIDNumber",
"HasPicture","Hobby","Home2TelephoneNumber",
"HomeAddress","HomeAddressCity",
				"HomeAddressCountry","HomeAddressPostalCode",
"HomeAddressPostOfficeBox","HomeAddressState",
"HomeAddressStreet","HomeFaxNumber",
"HomeTelephoneNumber","IMAddress","Importance",
"Initials","InternetFreeBusyAddress","IsConflict",
"ISDNNumber","ItemProperties","JobTitle",
"Journal","Language","LastFirstAndSuffix",
"LastFirstNoSpace","LastFirstNoSpaceAndSuffix",
"LastFirstNoSpaceCompany",
				"LastFirstSpaceOnly","LastFirstSpaceOnlyCompany",
"LastModificationTime","LastName",
"LastNameAndFirstName","Links","MailingAddress",
"MailingAddressCity","MailingAddressCountry",
"MailingAddressPostalCode","MailingAddressPostOfficeBox",
"MailingAddressState","MailingAddressStreet",
"ManagerName","MarkForDownload","MessageClass",
"MiddleName","Mileage","MobileTelephoneNumber",
"NetMeetingAlias",
				"NetMeetingServer","NickName","NoAging",
"OfficeLocation","OrganizationalIDNumber","OtherAddress",
"OtherAddressCity","OtherAddressCountry",
"OtherAddressPostalCode","OtherAddressPostOfficeBox",
"OtherAddressState","OtherAddressStreet",
"OtherFaxNumber","OtherTelephoneNumber",
"OutlookInternalVersion","OutlookVersion",
"PagerNumber","Parent","PersonalHomePage",
"PrimaryTelephoneNumber","Profession",
				"RadioTelephoneNumber","ReferredBy",
"Saved","SelectedMailingAddress","Sensitivity",
"Session","Size","Spouse","Subject",
"Suffix","TelexNumber","Title",
"TTYTDDTelephoneNumber","UnRead","User1",
"User2","User3","User4","UserCertificate",
"UserProperties","WebPage","YomiCompanyName",
"YomiFirstName","YomiLastName"
			};

		private Microsoft.Office.Interop.Outlook.Application application;

		public CompareThread(Microsoft.Office.Interop.Outlook.Application application)
		{
			this.application=application;
			//Thread thread = new Thread(new ThreadStart(Run));
			//thread.Start();
			Run();
		}

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
					}
					else
					{
						list = new ArrayList();
						cache.Add(value,list);
					}
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
							if ((testprop!=knownprop))
							{
								exact=false;
								break;
							}
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

				MessageBox.Show("Found "+redundant.Count+" exact duplicates");
				CompareResults results = new CompareResults(duplicateList,cache,props);
				results.Show();
			}
		}
	}
}
