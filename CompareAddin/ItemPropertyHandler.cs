/*
 * $LastChangedBy$
 * $HeadURL$
 * $Date$
 * $Revision$
 */

using System;
using Microsoft.Office.Interop.Outlook;

namespace CompareAddin
{
	public class ItemPropertyHandler
	{
		private OlItemType type;
		private string index;

		public ItemPropertyHandler(OlItemType type, string index)
		{
			this.type=type;
			this.index=index;
		}

		public UserProperties FetchUserProperties(object obj)
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

		private string Normalise(string value)
		{
			return value.Trim().ToLower();
		}

		private static string NameProperty(OlItemType type)
		{
			if (type==OlItemType.olContactItem)
			{
				return "FileAs";
			}
			else
			{
				return null;
			}
		}

		public MAPIFolder FetchFolder(object obj)
		{
			if (obj is ContactItem)
			{
				ContactItem item = (ContactItem)obj;
				return (MAPIFolder)item.Parent;
			}
			else
			{
				return null;
			}
		}

		public string FetchNameProperty(object obj)
		{
			UserProperties props = FetchUserProperties(obj);
			if (props!=null)
			{
				UserProperty prop = props.Find(NameProperty(type),false);
				if (prop!=null)
				{
					return (string)prop.Value;
				}
			}
			return null;
		}

		public string FetchIndexProperty(object obj)
		{
			UserProperties props = FetchUserProperties(obj);
			if (props!=null)
			{
				UserProperty prop = props.Find(index,false);
				if (prop!=null)
				{
					return Normalise((string)prop.Value);
				}
			}
			return null;
		}

		public bool IsCorrectType(object obj)
		{
			if ((type==OlItemType.olContactItem)&&(obj is ContactItem))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
	}
}
