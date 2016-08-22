﻿using System;
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
        private string[] lines = null;
		private readonly Missing MISS=Missing.Value;
		/// <summary>Contains a reference to the hosting application.</summary>
		private Microsoft.Office.Interop.Excel.Application m_XlApplication=null;
		/// <summary>Contains a reference to the active workbook.</summary>
		private Workbook m_Workbook=null;
		private bool m_ToolBarVisible=true;
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
	#endregion Properties

	#region Events

		private void OnWebBrowserExcelNavigated(object sender,WebBrowserNavigatedEventArgs e) {
			AttachApplication();
		}

	#endregion Events

	#region Methods
		public void OpenFile(string filename,string [] _lines) {
            lines = _lines;
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
                GetDiff();
				// Create the Excel.Application object
				m_XlApplication=(Microsoft.Office.Interop.Excel.Application)m_Workbook.Application;
			} catch(Exception ex) {
				MessageBox.Show(ex.Message);
				return;
			}
		}
        
		public Worksheet FindExcelWorksheet(string sheetname) {
            if (m_Workbook.Worksheets == null) return null;
			Worksheet sheet=null;
			// Step through the worksheet collection and see if the sheet is available. If found return true;
            for (int isheet = 1; isheet <= m_Workbook.Worksheets.Count; isheet++)
            {
                sheet = (Worksheet)m_Workbook.Worksheets.get_Item((object)isheet);
                if (sheet.Name.Equals(sheetname)) { return sheet; }
			}
			return null;
		}
	#endregion Methods

        internal void Close()
        {
            try
            {
                // Quit Excel and clean up.
                if (m_Workbook != null)
                {
                    m_Workbook.Close(true, Missing.Value, Missing.Value);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject
                                            (m_Workbook);
                    m_Workbook = null;
                }
                if (m_XlApplication != null)
                {
                    m_XlApplication.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject
                                        (m_XlApplication);
                    m_XlApplication = null;
                    System.GC.Collect();
                }
            }
            catch
            {
                MessageBox.Show("Failed to close the application");
            }
        }


        private void GetDiff()
        {
            
            
            foreach (var line in lines)
            {
                if (line.Length > 20)
                {
                    string _old = "";
                    string _new = "";
                    string sheet = line.Substring(18, line.IndexOf("!", 18) - 18);
                    string range = line.Substring(line.IndexOf("!", 18) + 1, line.IndexOf(" ", 18) - line.IndexOf("!", 18) -1);

                    Worksheet wSheet = FindExcelWorksheet(sheet);
                    if (wSheet == null)
                    {
                        wSheet = m_Workbook.Worksheets.Add(Type.Missing, m_Workbook.Worksheets[m_Workbook.Worksheets.Count]);
                        wSheet.Name = sheet;
                    }
                    
                    if (line.Substring(14, 3) == "   ")
                    {
                        _old = line.Substring(32, line.IndexOf("' v/s '") - 32);
                        _new = line.Substring(line.IndexOf("' v/s '") + 7, line.Length - line.IndexOf("' v/s '") - 9);
                        wSheet.Range[range].AddComment("Changed Old Value : " + _old);
                        wSheet.Range[range].Value2 = _new;
                    }
                    else if (line.Substring(14, 3) == "WB1")
                    {
                        _old = line.Substring(32, line.Length - 34);
                        wSheet.Range[range].AddComment("Deleted Value : " + _old);
                        wSheet.Range[range].Value2 = "";
                    }
                    else if (line.Substring(14, 3) == "WB2")
                    {
                        _new = line.Substring(32, line.Length - 34);
                        wSheet.Range[range].AddComment("Added");
                        wSheet.Range[range].Value2=_new;
                    }
                }
            }
        }

    }
}
