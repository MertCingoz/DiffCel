using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace EmbeddedExcel
{
    public partial class ExcelWrapper : UserControl
    {

        [DllImport("ole32.dll")]
        static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable pprot);
        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);

        #region Fields
        private readonly Missing MISS = Missing.Value;
        /// <summary>Contains a reference to the hosting application.</summary>
        private Microsoft.Office.Interop.Excel.Application m_XlApplication = null;
        /// <summary>Contains a reference to the active workbook.</summary>
        private Workbook m_Workbook = null;
        /// <summary>Contains the path to the workbook file.</summary>
        private string m_ExcelFileName = string.Empty;
        #endregion Fields

        #region Construction
        public ExcelWrapper()
        {
            InitializeComponent();
        }
        #endregion Construction

        #region Properties
        [Browsable(false)]
        public Workbook Workbook
        {
            get { return m_Workbook; }
        }
        #endregion Properties

        #region Events

        private void OnWebBrowserExcelNavigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            if (e.Url.ToString().Contains("Temp"))
                AttachApplication();
        }

        private void WebBrowserExcel_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            this.Visible = true;
        }
        #endregion Events

        #region Methods
        public void OpenFile(string filename)
        {
            // Check the file exists
            if (!System.IO.File.Exists(filename)) throw new Exception();
            m_ExcelFileName = filename.Replace("\\", "/");
            // Load the workbook in the WebBrowser control
            WebBrowserExcel.Navigate(filename, false);
        }

        public Workbook GetActiveWorkbook(string xlfile)
        {
            IRunningObjectTable prot = null;
            IEnumMoniker pmonkenum = null;
            try
            {
                IntPtr pfetched = IntPtr.Zero;
                // Query the running object table (ROT)
                if (GetRunningObjectTable(0, out prot) != 0 || prot == null) return null;
                prot.EnumRunning(out pmonkenum);
                pmonkenum.Reset();
                IMoniker[] monikers = new IMoniker[1];
                while (pmonkenum.Next(1, monikers, pfetched) == 0)
                {
                    IBindCtx pctx; string filepathname;
                    CreateBindCtx(0, out pctx);
                    // Get the name of the file
                    monikers[0].GetDisplayName(pctx, null, out filepathname);
                    // Clean up
                    Marshal.ReleaseComObject(pctx);
                    // Search for the workbook
                    filepathname=filepathname.Replace("\\", "/");
                    if (filepathname.IndexOf(xlfile) != -1)
                    {
                        object roval;
                        // Get a handle on the workbook
                        prot.GetObject(monikers[0], out roval);
                        return roval as Workbook;
                    }
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                // Clean up
                if (prot != null) Marshal.ReleaseComObject(prot);
                if (pmonkenum != null) Marshal.ReleaseComObject(pmonkenum);
            }
            return null;
        }

        private void AttachApplication()
        {
            try
            {
                if (m_ExcelFileName == null || m_ExcelFileName.Length == 0) return;
                // Creation of the workbook object
                if ((m_Workbook = GetActiveWorkbook(m_ExcelFileName)) == null) return;
                //GetDiff();
                // Create the Excel.Application object
                m_XlApplication = m_Workbook.Application;
                //CommandBar m_StandardCommandBar = m_XlApplication.CommandBars["Standart"];
                //m_StandardCommandBar.Position = MsoBarPosition.msoBarTop;
                //m_StandardCommandBar.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GetDiff()
        {
            List<string> sheets = new List<string>();
            foreach (Worksheet sheet in m_Workbook.Worksheets)
                sheets.Add(sheet.Name);
            foreach (Cell cell in Form1.cells)
                if (!sheets.Contains(cell.Sheet))
                {
                    m_Workbook.Worksheets.Add(Type.Missing, m_Workbook.Worksheets[m_Workbook.Worksheets.Count]).Name = cell.Sheet;
                    sheets.Add(cell.Sheet);
                }

            foreach (Cell cell in Form1.cells)
            {
                int index = sheets.IndexOf(cell.Sheet) + 1;
                m_Workbook.Worksheets[index].Range[cell.Adress].AddComment("");
            }
        }

        internal void FocusCell(Cell cell)
        {
            try
            {
                m_Workbook.Worksheets[cell.Sheet].Activate();
                m_Workbook.Worksheets[cell.Sheet].Select();
                m_Workbook.Worksheets[cell.Sheet].Range[cell.Adress].Select();
            }
            catch
            {
                
            }
        }
        #endregion Methods

    }
}

