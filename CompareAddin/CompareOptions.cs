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
	public class CompareOptions : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtFolder1;
		private System.Windows.Forms.TextBox txtFolder2;
		private System.Windows.Forms.Button btnFolder1;
		private System.Windows.Forms.Button btnFolder2;
		private System.Windows.Forms.RadioButton radioSingleFolder;
		private System.Windows.Forms.RadioButton radioMultiFolder;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.ComboBox cmbType;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ComboBox cmbField;
		private System.Windows.Forms.Button btnCompare;
		private System.Windows.Forms.Button btnCancel;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		private NameSpace ns;
		private MAPIFolder folder1,folder2;

		public CompareOptions(NameSpace ns)
		{
			InitializeComponent();
			this.ns=ns;
			cmbType.SelectedIndex=0;
			cmbField.SelectedIndex=1;
		}

		public bool CompareMultipleFolders
		{
			get
			{
				return radioMultiFolder.Checked;
			}
		}

		public string Field
		{
			get
			{
				return (string)cmbField.SelectedItem;
			}
		}

		public MAPIFolder Folder1
		{
			get
			{
				return folder1;
			}
		}

		public MAPIFolder Folder2
		{
			get
			{
				return folder2;
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
			this.txtFolder1 = new System.Windows.Forms.TextBox();
			this.txtFolder2 = new System.Windows.Forms.TextBox();
			this.btnFolder1 = new System.Windows.Forms.Button();
			this.btnFolder2 = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.radioSingleFolder = new System.Windows.Forms.RadioButton();
			this.radioMultiFolder = new System.Windows.Forms.RadioButton();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbType = new System.Windows.Forms.ComboBox();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbField = new System.Windows.Forms.ComboBox();
			this.btnCompare = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// txtFolder1
			// 
			this.txtFolder1.Enabled = false;
			this.txtFolder1.Location = new System.Drawing.Point(128, 104);
			this.txtFolder1.Name = "txtFolder1";
			this.txtFolder1.Size = new System.Drawing.Size(200, 20);
			this.txtFolder1.TabIndex = 0;
			this.txtFolder1.Text = "";
			// 
			// txtFolder2
			// 
			this.txtFolder2.Enabled = false;
			this.txtFolder2.Location = new System.Drawing.Point(128, 144);
			this.txtFolder2.Name = "txtFolder2";
			this.txtFolder2.Size = new System.Drawing.Size(200, 20);
			this.txtFolder2.TabIndex = 1;
			this.txtFolder2.Text = "";
			// 
			// btnFolder1
			// 
			this.btnFolder1.Location = new System.Drawing.Point(344, 104);
			this.btnFolder1.Name = "btnFolder1";
			this.btnFolder1.TabIndex = 2;
			this.btnFolder1.Text = "Browse...";
			this.btnFolder1.Click += new System.EventHandler(this.btnFolder1_Click);
			// 
			// btnFolder2
			// 
			this.btnFolder2.Location = new System.Drawing.Point(344, 144);
			this.btnFolder2.Name = "btnFolder2";
			this.btnFolder2.TabIndex = 3;
			this.btnFolder2.Text = "Browse...";
			this.btnFolder2.Click += new System.EventHandler(this.btnFolder2_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(24, 104);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(100, 16);
			this.label1.TabIndex = 4;
			this.label1.Text = "First folder:";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(24, 144);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(100, 16);
			this.label2.TabIndex = 5;
			this.label2.Text = "Second folder:";
			// 
			// radioSingleFolder
			// 
			this.radioSingleFolder.Location = new System.Drawing.Point(32, 16);
			this.radioSingleFolder.Name = "radioSingleFolder";
			this.radioSingleFolder.Size = new System.Drawing.Size(208, 24);
			this.radioSingleFolder.TabIndex = 6;
			this.radioSingleFolder.Text = "Find duplicates within a single folder";
			// 
			// radioMultiFolder
			// 
			this.radioMultiFolder.Checked = true;
			this.radioMultiFolder.Location = new System.Drawing.Point(32, 48);
			this.radioMultiFolder.Name = "radioMultiFolder";
			this.radioMultiFolder.Size = new System.Drawing.Size(216, 24);
			this.radioMultiFolder.TabIndex = 7;
			this.radioMultiFolder.TabStop = true;
			this.radioMultiFolder.Text = "Find duplicates between two folders";
			this.radioMultiFolder.CheckedChanged += new System.EventHandler(this.radioMultiFolder_CheckedChanged);
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(24, 192);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(72, 24);
			this.label3.TabIndex = 8;
			this.label3.Text = "Item type:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbType
			// 
			this.cmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbType.Items.AddRange(new object[] {
																								 "Contacts"});
			this.cmbType.Location = new System.Drawing.Point(128, 192);
			this.cmbType.Name = "cmbType";
			this.cmbType.Size = new System.Drawing.Size(200, 21);
			this.cmbType.TabIndex = 9;
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(24, 232);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 23);
			this.label4.TabIndex = 10;
			this.label4.Text = "Comparison field:";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbField
			// 
			this.cmbField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbField.Items.AddRange(new object[] {
																									"Email1Address",
																									"FileAs",
																									"FullName"});
			this.cmbField.Location = new System.Drawing.Point(128, 232);
			this.cmbField.Name = "cmbField";
			this.cmbField.Size = new System.Drawing.Size(200, 21);
			this.cmbField.TabIndex = 11;
			// 
			// btnCompare
			// 
			this.btnCompare.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnCompare.Enabled = false;
			this.btnCompare.Location = new System.Drawing.Point(96, 280);
			this.btnCompare.Name = "btnCompare";
			this.btnCompare.Size = new System.Drawing.Size(104, 23);
			this.btnCompare.TabIndex = 12;
			this.btnCompare.Text = "Run Comparison";
			// 
			// btnCancel
			// 
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(248, 280);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(104, 23);
			this.btnCancel.TabIndex = 13;
			this.btnCancel.Text = "Cancel";
			// 
			// CompareOptions
			// 
			this.AcceptButton = this.btnCompare;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btnCancel;
			this.ClientSize = new System.Drawing.Size(450, 325);
			this.ControlBox = false;
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnCompare);
			this.Controls.Add(this.cmbField);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.cmbType);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.radioMultiFolder);
			this.Controls.Add(this.radioSingleFolder);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnFolder2);
			this.Controls.Add(this.btnFolder1);
			this.Controls.Add(this.txtFolder2);
			this.Controls.Add(this.txtFolder1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "CompareOptions";
			this.Text = "CompareOptions";
			this.ResumeLayout(false);

		}
		#endregion

		private void radioMultiFolder_CheckedChanged(object sender, System.EventArgs e)
		{
			if (radioMultiFolder.Checked)
			{
				btnFolder2.Enabled=true;
				if (folder2==null)
				{
					btnCompare.Enabled=false;
				}
			}
			else
			{
				btnFolder2.Enabled=false;
				if (folder1!=null)
				{
					btnCompare.Enabled=true;
				}
			}
		}

		private void btnFolder1_Click(object sender, System.EventArgs e)
		{
			MAPIFolder folder = ns.PickFolder();
			if (folder!=null)
			{
				folder1=folder;
				txtFolder1.Text=folder.FullFolderPath;
			}
			if ((radioSingleFolder.Checked)||(folder2!=null))
			{
				btnCompare.Enabled=true;
			}
		}

		private void btnFolder2_Click(object sender, System.EventArgs e)
		{
			MAPIFolder folder = ns.PickFolder();
			if (folder!=null)
			{
				folder2=folder;
				txtFolder1.Text=folder.FullFolderPath;
				if (folder1!=null)
				{
					btnCompare.Enabled=true;
				}
			}
		}
	}
}
