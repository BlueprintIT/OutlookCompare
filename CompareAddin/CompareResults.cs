/*
 * $LastChangedBy$
 * $HeadURL$
 * $Date$
 * $Revision$
 */

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace CompareAddin
{
	public class CompareResults : System.Windows.Forms.Form
	{
		private System.Windows.Forms.ListBox listItems;
		private System.Windows.Forms.ListBox listDuplicates;
		private System.Windows.Forms.Button btnOK;
		private System.ComponentModel.Container components = null;

		private IDictionary cache;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Splitter splitter1;
		private System.Windows.Forms.Panel panel2;
		private ItemPropertyHandler props;

		private class OutlookItem
		{
			private object item;
			private ItemPropertyHandler props;

			public OutlookItem(object item, ItemPropertyHandler props)
			{
				this.item=item;
				this.props=props;
			}

			public void Display()
			{
				if (item is ContactItem)
				{
					ContactItem oitem = (ContactItem)item;
					oitem.Display(false);
				}
			}

			public override string ToString()
			{
				return props.FetchFolder(item).FullFolderPath+" ("+props.FetchNameProperty(item)+")";
			}
		}

		public CompareResults(IList duplicates, IDictionary cache, ItemPropertyHandler props)
		{
			InitializeComponent();

			this.cache=cache;
			this.props=props;

			foreach (string key in duplicates)
			{
				listItems.Items.Add(key);
			}
		}

		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.listItems = new System.Windows.Forms.ListBox();
			this.listDuplicates = new System.Windows.Forms.ListBox();
			this.btnOK = new System.Windows.Forms.Button();
			this.panel1 = new System.Windows.Forms.Panel();
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.panel2 = new System.Windows.Forms.Panel();
			this.panel1.SuspendLayout();
			this.panel2.SuspendLayout();
			this.SuspendLayout();
			// 
			// listItems
			// 
			this.listItems.Dock = System.Windows.Forms.DockStyle.Left;
			this.listItems.Location = new System.Drawing.Point(0, 0);
			this.listItems.Name = "listItems";
			this.listItems.Size = new System.Drawing.Size(136, 316);
			this.listItems.TabIndex = 0;
			this.listItems.SelectedIndexChanged += new System.EventHandler(this.listItems_SelectedIndexChanged);
			// 
			// listDuplicates
			// 
			this.listDuplicates.Dock = System.Windows.Forms.DockStyle.Fill;
			this.listDuplicates.Location = new System.Drawing.Point(139, 0);
			this.listDuplicates.Name = "listDuplicates";
			this.listDuplicates.Size = new System.Drawing.Size(405, 316);
			this.listDuplicates.TabIndex = 1;
			this.listDuplicates.DoubleClick += new System.EventHandler(this.listDuplicates_DoubleClick);
			// 
			// btnOK
			// 
			this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnOK.Location = new System.Drawing.Point(456, 16);
			this.btnOK.Name = "btnOK";
			this.btnOK.TabIndex = 2;
			this.btnOK.Text = "OK";
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.listDuplicates);
			this.panel1.Controls.Add(this.splitter1);
			this.panel1.Controls.Add(this.listItems);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(544, 323);
			this.panel1.TabIndex = 3;
			// 
			// splitter1
			// 
			this.splitter1.Location = new System.Drawing.Point(136, 0);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(3, 323);
			this.splitter1.TabIndex = 1;
			this.splitter1.TabStop = false;
			// 
			// panel2
			// 
			this.panel2.Controls.Add(this.btnOK);
			this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.panel2.Location = new System.Drawing.Point(0, 323);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(544, 56);
			this.panel2.TabIndex = 4;
			// 
			// CompareResults
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(544, 379);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.panel2);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "CompareResults";
			this.Text = "Results of Comparison";
			this.panel1.ResumeLayout(false);
			this.panel2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void listItems_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			string key = (string)listItems.SelectedItem;
			listDuplicates.Items.Clear();
			IList items = (IList)cache[key];
			foreach (object obj in items)
			{
				listDuplicates.Items.Add(new OutlookItem(obj,props));
			}
		}

		private void listDuplicates_DoubleClick(object sender, System.EventArgs e)
		{
			((OutlookItem)listDuplicates.SelectedItem).Display();
		}
	}
}
