using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace EmbeddedExcel
{
    public partial class Form1 : Form
    {
        public static List<Cell> cells;
        private static string[] excelFormats = { "xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xls", ".xlt", ".xml", ".xlam", ".xlw" };
        private Process cmd = new Process();
        private string path;
        private string extension;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            excelWrapper.Close();
            ListDirectory(treeView1, AppDomain.CurrentDomain.BaseDirectory);
            cells = new List<Cell>();
        }

        private void ListDirectory(TreeView treeView, string path)
        {
            treeView.Nodes.Clear();
            var rootDirectoryInfo = new DirectoryInfo(path);
            treeView.Nodes.Add(CreateDirectoryNode(rootDirectoryInfo));
            treeView.ExpandAll();
        }

        private static TreeNode CreateDirectoryNode(DirectoryInfo directoryInfo)
        {
            var directoryNode = new TreeNode(directoryInfo.Name);
            foreach (var directory in directoryInfo.GetDirectories())
                if (CreateDirectoryNode(directory) != null)
                    directoryNode.Nodes.Add(CreateDirectoryNode(directory));
            foreach (var file in directoryInfo.GetFiles())
                if (excelFormats.Any(file.Name.Contains) && !file.Name.Contains("Temp"))
                    directoryNode.Nodes.Add(new TreeNode(file.Name));
            if (directoryNode.Nodes.Count == 0)
                return null;
            return directoryNode;
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Nodes.Count == 0)
            {
                path = "";
                TreeNode temp = e.Node;
                while (temp.Parent != null)
                {
                    path = "\\" + temp.Text + path;
                    temp = temp.Parent;
                }
                path = path.Substring(1, path.Length - 1);

                cmd.StartInfo.FileName = "cmd.exe";
                cmd.StartInfo.RedirectStandardInput = true;
                cmd.StartInfo.RedirectStandardOutput = true;
                cmd.StartInfo.CreateNoWindow = true;
                cmd.StartInfo.UseShellExecute = false;
                cmd.Start();
                cmd.StandardInput.WriteLine("git log --pretty=format:\"%h|%an|%s\" \"" + path + "\" >raw.txt");
                cmd.StandardInput.Close();
                cmd.WaitForExit();

                listView1.Items.Clear();
                string[] lines = System.IO.File.ReadAllLines(@"raw.txt");
                foreach (string line in lines)
                {
                    string[] objects = line.Split('|');
                    listView1.Items.Add(objects[0]);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(objects[1]);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(objects[2]);
                }
                excelWrapper.Close();
            }
        }
        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected)
            {
                try
                {
                    var excelProcesses = Process.GetProcessesByName("excel");
                    foreach (var process in excelProcesses)
                        if (process.MainWindowTitle == "" && DialogResult.Yes == MessageBox.Show("Excel application being used by another process.\nEnsure that your local works are SAVED.\n\nDo you want to terminate EXCEL processes ?", "Opps", MessageBoxButtons.YesNoCancel))
                        {

                            process.Kill();
                            Cursor.Current = Cursors.WaitCursor;
                            process.WaitForExit();
                        }
                    
                    string[] files = System.IO.Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "Temp*", System.IO.SearchOption.TopDirectoryOnly);
                    foreach (var file in files)
                        File.Delete(file);
                    string dir = path.Substring(0, path.LastIndexOf("\\") + 1);
                    extension = path.Substring(path.LastIndexOf("."));
                    cmd.Start();
                    cmd.StandardInput.WriteLine("git diff " + e.Item.SubItems[0].Text + " \"" + path + "\" >raw2.txt");
                    cmd.StandardInput.WriteLine("git cat-file -p " + e.Item.SubItems[0].Text + ":\"" + path.Replace('\\', '/') + "\" > Temp" + extension);
                    cmd.StandardInput.Close();
                    cmd.WaitForExit();
                    listView2.Items.Clear();
                    excelWrapper.Visible = false;
                    GetDiff();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void GetDiff()
        {
            string raw = System.IO.File.ReadAllText(@"raw2.txt");
            if (raw.Length == 0) return;
            raw = raw.Substring(0, raw.IndexOf("----------------- DIFF -------------------") - 1);
            string[] lines = raw.Split('\n');
            cells.Clear();
            foreach (var line in lines)
            {
                Cell cell = new Cell();
                cell.Sheet = line.Substring(18, line.IndexOf("!", 18) - 18);
                cell.Adress = line.Substring(line.IndexOf("!", 18) + 1, line.IndexOf(" ", 18) - line.IndexOf("!", 18) - 1);
                if (line.Substring(14, 3) == "   ")
                {
                    cell.OldValue = line.Substring(32, line.IndexOf("' v/s '") - 32);
                    cell.NewValue = line.Substring(line.IndexOf("' v/s '") + 7, line.Length - line.IndexOf("' v/s '") - 9);
                    cell.Operation = "Changed";
                }
                else if (line.Substring(14, 3) == "WB1")
                {
                    cell.OldValue = line.Substring(32, line.Length - 34);
                    cell.Operation = "Deleted";
                }
                else if (line.Substring(14, 3) == "WB2")
                {
                    cell.NewValue = line.Substring(32, line.Length - 34);
                    cell.Operation = "Added";
                }
                cells.Add(cell);
                listView2.Items.Add(cell.Operation);
                listView2.Items[listView2.Items.Count - 1].SubItems.Add(cell.Sheet);
                listView2.Items[listView2.Items.Count - 1].SubItems.Add(cell.Adress);
                listView2.Items[listView2.Items.Count - 1].SubItems.Add(cell.OldValue);
                listView2.Items[listView2.Items.Count - 1].SubItems.Add(cell.NewValue);
            }
        }

        
        private void listView2_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            Thread thread = new Thread(() => excelWrapper.OpenFile(AppDomain.CurrentDomain.BaseDirectory + "Temp" + extension, cells[e.ItemIndex]));
            thread.Start();
        }

        private void listView_MouseMove(object sender, MouseEventArgs e)
        {
            ListView lv = (ListView)sender;
            var hit = lv.HitTest(e.Location);
            if (hit.SubItem != null && hit.SubItem == hit.Item.SubItems[0])
                lv.Cursor = Cursors.Hand;
            else
                lv.Cursor = Cursors.Default;
        }


    }
}