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
	public class ProgressDialog : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label;
		private System.Windows.Forms.ProgressBar progressBar;
		private System.ComponentModel.Container components = null;

		public ProgressDialog(string title, string text)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			label.Text=text;
			Text=title;
		}

		public int Value
		{
			get
			{
				return progressBar.Value;
			}

			set
			{
				progressBar.Value=value;
				Redraw();
			}
		}

		public int Maximum
		{
			get
			{
				return progressBar.Maximum;
			}

			set
			{
				progressBar.Maximum=value;
				Redraw();
			}
		}

		private void Redraw()
		{
			Refresh();
			System.Threading.Thread.Sleep(0);
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
			this.label = new System.Windows.Forms.Label();
			this.progressBar = new System.Windows.Forms.ProgressBar();
			this.SuspendLayout();
			// 
			// label
			// 
			this.label.Location = new System.Drawing.Point(8, 16);
			this.label.Name = "label";
			this.label.Size = new System.Drawing.Size(320, 23);
			this.label.TabIndex = 0;
			this.label.Text = "label";
			this.label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// progressBar
			// 
			this.progressBar.Location = new System.Drawing.Point(16, 48);
			this.progressBar.Name = "progressBar";
			this.progressBar.Size = new System.Drawing.Size(304, 16);
			this.progressBar.TabIndex = 1;
			// 
			// ProgressDialog
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(338, 77);
			this.ControlBox = false;
			this.Controls.Add(this.progressBar);
			this.Controls.Add(this.label);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "ProgressDialog";
			this.Text = "ProgressDialog";
			this.ResumeLayout(false);

		}
		#endregion
	}
}
