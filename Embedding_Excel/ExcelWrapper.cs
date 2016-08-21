using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Office=Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;


namespace EmbeddedExcel
{
	public partial class ExcelWrapper : UserControl {

		[DllImport("ole32.dll")]
		static extern int GetRunningObjectTable(uint reserved,out IRunningObjectTable pprot);
		[DllImport("ole32.dll")]
		static extern int CreateBindCtx(uint reserved,out IBindCtx pctx);

	#region Fields
		private readonly Missing MISS=Missing.Value;
		/// <summary>Contains a reference to the hosting application.</summary>
		private Microsoft.Office.Interop.Excel.Application m_XlApplication=null;
		/// <summary>Contains a reference to the active workbook.</summary>
		private Workbook m_Workbook=null;
		private bool m_ToolBarVisible=true;
		private Office.CommandBar m_StandardCommandBar=null;
		/// <summary>Contains the path to the workbook file.</summary>
		private string m_ExcelFileName=string.Empty;
	#endregion Fields

	#region Construction
		public ExcelWrapper() {
			InitializeComponent();
		}
	#endregion Construction

	#region Properties
		[Browsable(false)]
		public Workbook Workbook {
			get { return m_Workbook; }
		}

		[Browsable(true),Category("Appearance")]
		public bool ToolBarVisible {
			get { return m_ToolBarVisible; }
			set {
				if(m_ToolBarVisible==value) return;
				m_ToolBarVisible=value;
				if(m_XlApplication!=null) OnToolBarVisibleChanged();
			}
		}
	#endregion Properties

	#region Events
		private void OnToolBarVisibleChanged() {
			try {
				m_StandardCommandBar.Visible=m_ToolBarVisible;
			} catch { }
		}

		private void OnWebBrowserExcelNavigated(object sender,WebBrowserNavigatedEventArgs e) {
			AttachApplication();
		}

		//private void OnOpenClick(Office.CommandBarButton Ctrl,ref bool Cancel) {
		//    if(this.OpenExcelFileDialog.ShowDialog()==DialogResult.OK) {
		//        OpenFile(OpenExcelFileDialog.FileName);
		//    }
		//}

		//void OnNewClick(Office.CommandBarButton Ctrl,ref bool Cancel) {
		//    throw new Exception("The method or operation is not implemented.");
		//}
	#endregion Events

	#region Methods
		public void OpenFile(string filename) {
			// Check the file exists
			if(!System.IO.File.Exists(filename)) throw new Exception();
			m_ExcelFileName=filename.Replace("\\","/");
			// Load the workbook in the WebBrowser control
            WebBrowserExcel.Navigate(filename, false);
		}

		public Workbook GetActiveWorkbook(string xlfile) {
			IRunningObjectTable prot=null;
			IEnumMoniker pmonkenum=null;
			try {
				IntPtr pfetched=IntPtr.Zero;
				// Query the running object table (ROT)
				if(GetRunningObjectTable(0,out prot)!=0||prot==null) return null;
				prot.EnumRunning(out pmonkenum);
				pmonkenum.Reset();
				IMoniker[] monikers=new IMoniker[1];
				while(pmonkenum.Next(1,monikers,pfetched)==0) {
					IBindCtx pctx; string filepathname;
					CreateBindCtx(0,out pctx);
					// Get the name of the file
					monikers[0].GetDisplayName(pctx,null,out filepathname);
					// Clean up
					Marshal.ReleaseComObject(pctx);
					// Search for the workbook
					if(filepathname.IndexOf(xlfile)!=-1) {
						object roval;
						// Get a handle on the workbook
						prot.GetObject(monikers[0],out roval);
                        return roval as Workbook;
					}
				}
			} catch {
				return null;
			} finally {
				// Clean up
				if(prot!=null) Marshal.ReleaseComObject(prot);
				if(pmonkenum!=null) Marshal.ReleaseComObject(pmonkenum);
			}
			return null;
		}

		private void AttachApplication() {
			try {
				if(m_ExcelFileName==null||m_ExcelFileName.Length==0) return;
				// Creation of the workbook object
				if((m_Workbook=GetActiveWorkbook(m_ExcelFileName))==null)return;
				// Create the Excel.Application object
				m_XlApplication=(Microsoft.Office.Interop.Excel.Application)m_Workbook.Application;
				// Creation of the standard toolbar
				m_StandardCommandBar=m_XlApplication.CommandBars["Standard"];
				m_StandardCommandBar.Position=Office.MsoBarPosition.msoBarTop;
				m_StandardCommandBar.Visible=m_ToolBarVisible;
				// Enable the OpenFile and New buttons
				foreach(Office.CommandBarControl control in m_StandardCommandBar.Controls) {
					string name = control.get_accName(Missing.Value);
					if(name.Equals("Nouveau")) ((Office.CommandBarButton)control).Enabled=false;
					if(name.Equals("Ouvrir")) ((Office.CommandBarButton)control).Enabled=false;
				}
			} catch {
				MessageBox.Show("Impossible de charger le fichier Excel");
				return;
			}
		}

		public Worksheet FindExcelWorksheet(string sheetname) {
			if(m_Workbook.Sheets==null) return null;
			Worksheet sheet=null;
			// Step through the worksheet collection and see if the sheet is available. If found return true;
			for(int isheet=1; isheet<=m_Workbook.Sheets.Count; isheet++) {
				sheet=(Worksheet)m_Workbook.Sheets.get_Item((object)isheet);
				if(sheet.Name.Equals(sheetname)) { sheet.Activate(); return sheet; }
			}
			return null;
		}
	#endregion Methods





    }
}

