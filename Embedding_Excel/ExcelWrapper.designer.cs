namespace EmbeddedExcel
{
	partial class ExcelWrapper
	{
		/// <summary>Required designer variable.</summary>
		private System.ComponentModel.IContainer components=null;

		/// <summary>Clean up any resources being used.</summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing) {
			if(disposing&&(components!=null)) {
				components.Dispose();
			}
			//this.WebBrowserExcel.Dispose();
			try {
				// Quit Excel and clean up.
				if(m_Workbook!=null) {
					m_Workbook.Close(true,MISS,MISS);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(m_Workbook);
					m_Workbook=null;
				}
				if(m_XlApplication!=null) {
					m_XlApplication.Quit();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(m_XlApplication);
					m_XlApplication=null;
					System.GC.Collect();
				}
			} catch {
				//MessageBox.Show("Impossible d'enregistrer la feuille 'Chi deux'");
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			this.WebBrowserExcel=new System.Windows.Forms.WebBrowser();
			this.OpenExcelFileDialog=new System.Windows.Forms.OpenFileDialog();
			this.SuspendLayout();
			// 
			// WebBrowserExcel
			// 
			this.WebBrowserExcel.Dock=System.Windows.Forms.DockStyle.Fill;
			this.WebBrowserExcel.Location=new System.Drawing.Point(0,0);
			this.WebBrowserExcel.MinimumSize=new System.Drawing.Size(20,20);
			this.WebBrowserExcel.Name="WebBrowserExcel";
			this.WebBrowserExcel.Size=new System.Drawing.Size(420,400);
			this.WebBrowserExcel.TabIndex=0;
			this.WebBrowserExcel.Navigated+=new System.Windows.Forms.WebBrowserNavigatedEventHandler(this.OnWebBrowserExcelNavigated);
			// 
			// OpenExcelFileDialog
			// 
			this.OpenExcelFileDialog.FileName="\"* Excel files | *.xls\"";
			// 
			// ExcelWrapper
			// 
			this.AutoScaleDimensions=new System.Drawing.SizeF(6F,13F);
			this.AutoScaleMode=System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.WebBrowserExcel);
			this.Name="ExcelWrapper";
			this.Size=new System.Drawing.Size(420,400);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.WebBrowser WebBrowserExcel;
		private System.Windows.Forms.OpenFileDialog OpenExcelFileDialog;
	}
}
