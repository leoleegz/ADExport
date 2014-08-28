using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using System.DirectoryServices;
using ActiveDs;

namespace ADExport
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class MainForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.TreeView tvwAD;
		private System.Windows.Forms.CheckBox chkIncludeSubOU;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.CheckBox chkExport;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnBrowse;
		private System.Windows.Forms.TextBox txtFileName;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnExport;
		private System.Windows.Forms.SaveFileDialog sfdExport;
		private System.Windows.Forms.GroupBox groupBox1;


		//private System.DirectoryServices.DirectoryEntry m_deRoot;
		private System.Windows.Forms.TreeNode m_tvRootNode;
		private System.Windows.Forms.TreeNode m_tvTempNode;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MainForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		/// <summary>
		/// 展开子节点
		/// </summary>
		/// <param name="tvParentNode"></param>
		private void ExpandSubNodes(TreeNode tvParentNode)
		{
			// 删除临时节点
			try
			{
				tvParentNode.Nodes.RemoveAt(0);
			}
			catch(Exception)
			{}

			using(DirectoryEntry deParent = new DirectoryEntry( tvParentNode.Tag.ToString(), null, null, AuthenticationTypes.Secure ))
			{
				foreach(DirectoryEntry deSub in deParent.Children )
				{
					// 只列出 OU 或者容器
					if( deSub.SchemaClassName == "organizationalUnit" || deSub.SchemaClassName == "container" )
					{
						TreeNode tvSubNode = new TreeNode( deSub.Properties["name"].Value.ToString() );
						tvSubNode.Tag = deSub.Path;
						tvParentNode.Nodes.Add( tvSubNode );
						tvSubNode.Nodes.Add( m_tvTempNode );
					}
				}
			}
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tvwAD = new System.Windows.Forms.TreeView();
			this.chkIncludeSubOU = new System.Windows.Forms.CheckBox();
			this.label1 = new System.Windows.Forms.Label();
			this.chkExport = new System.Windows.Forms.CheckBox();
			this.label2 = new System.Windows.Forms.Label();
			this.txtFileName = new System.Windows.Forms.TextBox();
			this.btnBrowse = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.btnExport = new System.Windows.Forms.Button();
			this.sfdExport = new System.Windows.Forms.SaveFileDialog();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.SuspendLayout();
			// 
			// tvwAD
			// 
			this.tvwAD.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tvwAD.ImageIndex = -1;
			this.tvwAD.Location = new System.Drawing.Point(8, 24);
			this.tvwAD.Name = "tvwAD";
			this.tvwAD.SelectedImageIndex = -1;
			this.tvwAD.Size = new System.Drawing.Size(416, 200);
			this.tvwAD.TabIndex = 0;
			this.tvwAD.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.tvwAD_BeforeExpand);
			// 
			// chkIncludeSubOU
			// 
			this.chkIncludeSubOU.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.chkIncludeSubOU.Checked = true;
			this.chkIncludeSubOU.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkIncludeSubOU.Location = new System.Drawing.Point(8, 232);
			this.chkIncludeSubOU.Name = "chkIncludeSubOU";
			this.chkIncludeSubOU.Size = new System.Drawing.Size(416, 16);
			this.chkIncludeSubOU.TabIndex = 1;
			this.chkIncludeSubOU.Text = "Include sub containers";
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(416, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Please select a container";
			// 
			// chkExport
			// 
			this.chkExport.Location = new System.Drawing.Point(8, 256);
			this.chkExport.Name = "chkExport";
			this.chkExport.Size = new System.Drawing.Size(424, 16);
			this.chkExport.TabIndex = 3;
			this.chkExport.Text = "Export object type";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 288);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 24);
			this.label2.TabIndex = 4;
			this.label2.Text = "File name:";
			// 
			// txtFileName
			// 
			this.txtFileName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtFileName.Location = new System.Drawing.Point(88, 288);
			this.txtFileName.Name = "txtFileName";
			this.txtFileName.Size = new System.Drawing.Size(248, 20);
			this.txtFileName.TabIndex = 5;
			this.txtFileName.Text = "";
			// 
			// btnBrowse
			// 
			this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnBrowse.Location = new System.Drawing.Point(344, 288);
			this.btnBrowse.Name = "btnBrowse";
			this.btnBrowse.Size = new System.Drawing.Size(80, 24);
			this.btnBrowse.TabIndex = 6;
			this.btnBrowse.Text = "Browse";
			this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
			// 
			// btnClose
			// 
			this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnClose.Location = new System.Drawing.Point(344, 320);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(80, 24);
			this.btnClose.TabIndex = 7;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// btnExport
			// 
			this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExport.Location = new System.Drawing.Point(256, 320);
			this.btnExport.Name = "btnExport";
			this.btnExport.Size = new System.Drawing.Size(80, 24);
			this.btnExport.TabIndex = 8;
			this.btnExport.Text = "Export";
			// 
			// sfdExport
			// 
			this.sfdExport.DefaultExt = "txt";
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(8, 276);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(416, 4);
			this.groupBox1.TabIndex = 9;
			this.groupBox1.TabStop = false;
			// 
			// MainForm
			// 
			this.AcceptButton = this.btnExport;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btnClose;
			this.ClientSize = new System.Drawing.Size(432, 357);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExport);
			this.Controls.Add(this.btnClose);
			this.Controls.Add(this.btnBrowse);
			this.Controls.Add(this.txtFileName);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.chkExport);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.chkIncludeSubOU);
			this.Controls.Add(this.tvwAD);
			this.Name = "MainForm";
			this.Text = "AD Objects Export";
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new MainForm());
		}

		private void MainForm_Load(object sender, System.EventArgs e)
		{
			string strDomainName = null;
			using( DirectoryEntry rootDE = new DirectoryEntry("LDAP://rootDSE", null, null, AuthenticationTypes.Secure) )
			{
				strDomainName = rootDE.Properties["defaultNamingContext"][0].ToString();
			}
			
			// 根节点
			ADSystemInfo objAD = new ADSystemInfoClass();
			m_tvRootNode = new TreeNode( objAD.DomainDNSName );
			m_tvRootNode.Tag = strDomainName;

			// 临时节点
			m_tvTempNode = new TreeNode( "Waiting for expanding ...");
			m_tvTempNode.Tag = "Temp";

			tvwAD.Nodes.Add( m_tvRootNode );
			//m_tvRootNode.Nodes.Add( m_tvTempNode );

			//ExpandSubNodes(m_tvRootNode);
			DirectoryEntry deRoot = new DirectoryEntry("LDAP://" + strDomainName);
			string[] aryProps = new string[] {"name"};
			using( DirectorySearcher deDS = new DirectorySearcher(deRoot, "(!(objectCategory=organizationalUnit)(objectCategory=container))", aryProps, SearchScope.OneLevel) )
			{
				SearchResultCollection deSRC = deDS.FindAll();
				for(int i = 0; i < deSRC.Count; i++)
				{
					TreeNode tvSubNode = new TreeNode( deSRC[i].Properties["name"][0].ToString() );
					tvSubNode.Tag = deSRC[i].Path;
					m_tvRootNode.Nodes.Add( tvSubNode );
					tvSubNode.Nodes.Add( m_tvTempNode );
				}
			}

			m_tvRootNode.Expand();
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			Application.Exit();
		}

		private void btnBrowse_Click(object sender, System.EventArgs e)
		{
			
			DialogResult dr = sfdExport.ShowDialog();
			if( dr == DialogResult.OK )
			{
				txtFileName.Text = sfdExport.FileName;
			}
		}

		private void tvwAD_BeforeExpand(object sender, System.Windows.Forms.TreeViewCancelEventArgs e)
		{
			// 如果节点含有 Temp 子节点，那么展开它
			if( e.Node.Nodes.Count == 1 )
			{
				if( e.Node.Nodes[0].Tag.ToString() == "Temp" )
				{
					ExpandSubNodes( e.Node );
				}
			}
		}
	}
}
