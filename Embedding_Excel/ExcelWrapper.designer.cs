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
            this.WebBrowserExcel = new System.Windows.Forms.WebBrowser();
            this.SuspendLayout();
            // 
            // WebBrowserExcel
            // 
            this.WebBrowserExcel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.WebBrowserExcel.Location = new System.Drawing.Point(0, 0);
            this.WebBrowserExcel.Margin = new System.Windows.Forms.Padding(4);
            this.WebBrowserExcel.MinimumSize = new System.Drawing.Size(27, 25);
            this.WebBrowserExcel.Name = "WebBrowserExcel";
            this.WebBrowserExcel.Size = new System.Drawing.Size(560, 492);
            this.WebBrowserExcel.TabIndex = 0;
            this.WebBrowserExcel.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.WebBrowserExcel_DocumentCompleted);
            this.WebBrowserExcel.Navigated += new System.Windows.Forms.WebBrowserNavigatedEventHandler(this.OnWebBrowserExcelNavigated);
            // 
            // ExcelWrapper
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.WebBrowserExcel);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ExcelWrapper";
            this.Size = new System.Drawing.Size(560, 492);
            this.ResumeLayout(false);

		}

		#endregion

        private System.Windows.Forms.WebBrowser WebBrowserExcel;

    }
}
