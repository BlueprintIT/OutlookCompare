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

namespace CompareAddin
{
	public class CompareResults : System.Windows.Forms.Form
	{
		private System.Windows.Forms.ListBox listItems;
		private System.Windows.Forms.ListBox listDuplicates;
		private System.Windows.Forms.Button btnOK;
		private System.ComponentModel.Container components = null;

		private IDictionary cache;
		private ItemPropertyHandler props;

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
			this.SuspendLayout();
			// 
			// listItems
			// 
			this.listItems.Location = new System.Drawing.Point(8, 8);
			this.listItems.Name = "listItems";
			this.listItems.Size = new System.Drawing.Size(240, 303);
			this.listItems.TabIndex = 0;
			this.listItems.SelectedValueChanged += new System.EventHandler(this.listItems_SelectedValueChanged);
			this.listItems.SelectedIndexChanged += new System.EventHandler(this.listItems_SelectedIndexChanged);
			// 
			// listDuplicates
			// 
			this.listDuplicates.Location = new System.Drawing.Point(256, 8);
			this.listDuplicates.Name = "listDuplicates";
			this.listDuplicates.Size = new System.Drawing.Size(320, 303);
			this.listDuplicates.TabIndex = 1;
			// 
			// btnOK
			// 
			this.btnOK.Location = new System.Drawing.Point(496, 328);
			this.btnOK.Name = "btnOK";
			this.btnOK.TabIndex = 2;
			this.btnOK.Text = "OK";
			// 
			// CompareResults
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(584, 363);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.listDuplicates);
			this.Controls.Add(this.listItems);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "CompareResults";
			this.Text = "Results of Comparison";
			this.ResumeLayout(false);

		}
		#endregion

		private void listItems_SelectedValueChanged(object sender, System.EventArgs e)
		{
		}

		private void listItems_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			string key = (string)listItems.SelectedItem;
			listDuplicates.Items.Clear();
			IList items = (IList)cache[key];
			foreach (object obj in items)
			{
				listDuplicates.Items.Add(props.FetchFolder(obj).FullFolderPath+" ("+props.FetchNameProperty(obj)+")");
			}
		}
	}
}
