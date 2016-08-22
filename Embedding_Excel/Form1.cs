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
        private FolderBrowserDialog gitFolder = new FolderBrowserDialog();
        private Process cmd = new Process();
        private string relativePath;
        private string extension;
        private bool select = true;
        private string lastCommit="";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            gitFolder.ShowNewFolderButton = false;
            gitFolder.ShowDialog();
            if (Directory.Exists(gitFolder.SelectedPath + "\\.git"))
            {
                ListDirectory(treeView1, gitFolder.SelectedPath);
                cells = new List<Cell>();
            }
            else
                Application.Exit();
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
                relativePath = "";
                TreeNode temp = e.Node;
                while (temp.Parent != null)
                {
                    relativePath = "\\" + temp.Text + relativePath;
                    temp = temp.Parent;
                }
                relativePath = relativePath.Substring(1, relativePath.Length - 1);

                cmd.StartInfo.FileName = "cmd.exe";
                cmd.StartInfo.RedirectStandardInput = true;
                cmd.StartInfo.RedirectStandardOutput = true;
                cmd.StartInfo.CreateNoWindow = true;
                cmd.StartInfo.UseShellExecute = false;
                cmd.Start();
                cmd.StandardInput.WriteLine("cd " + gitFolder.SelectedPath);
                cmd.StandardInput.WriteLine("git log --pretty=format:\"%h|%an|%s\" \""+ relativePath + "\" >commits.txt");
                cmd.StandardInput.Close();
                cmd.WaitForExit();

                listView1.Items.Clear();
                string[] lines = System.IO.File.ReadAllLines(gitFolder.SelectedPath + "\\commits.txt");
                if (File.Exists(gitFolder.SelectedPath + "\\commits.txt"))
                    File.Delete(gitFolder.SelectedPath + "\\commits.txt");
                foreach (string line in lines)
                {
                    string[] objects = line.Split('|');
                    listView1.Items.Add(objects[0]);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(objects[1]);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(objects[2]);
                }
            }
        }
        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected && e.Item.SubItems[0].Text!=lastCommit)
            {
                listView2.Items.Clear();
                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    excelWrapper.Dispose();
                    excelWrapper = new EmbeddedExcel.ExcelWrapper();
                    splitContainer1.Panel2.Controls.Add(this.excelWrapper);
                    excelWrapper.Dock = System.Windows.Forms.DockStyle.Fill;
                    excelWrapper.Location = new System.Drawing.Point(0, 0);
                    excelWrapper.Margin = new System.Windows.Forms.Padding(5);
                    excelWrapper.Name = "excelWrapper";
                    excelWrapper.Size = new System.Drawing.Size(599, 608);
                    excelWrapper.TabIndex = 7;
                    excelWrapper.Visible = false;

                    string[] files = System.IO.Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "Temp*", System.IO.SearchOption.TopDirectoryOnly);
                    foreach (var file in files)
                        File.Delete(file);

                    string dir = relativePath.Substring(0, relativePath.LastIndexOf("\\") + 1);
                    extension = relativePath.Substring(relativePath.LastIndexOf("."));
                    cmd.Start();
                    cmd.StandardInput.WriteLine("cd " + gitFolder.SelectedPath);
                    cmd.StandardInput.WriteLine("git diff " + e.Item.SubItems[0].Text + " \"" + gitFolder.SelectedPath + "\\" + relativePath + "\" >diff.txt");
                    cmd.StandardInput.WriteLine("git cat-file -p " + e.Item.SubItems[0].Text + ":\"" + relativePath.Replace('\\', '/') + "\" > Temp" + extension);
                    cmd.StandardInput.Close();
                    cmd.WaitForExit();
                    GetDiff();
                    select = true;
                    lastCommit = e.Item.SubItems[0].Text;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void GetDiff()
        {
            string raw = System.IO.File.ReadAllText(gitFolder.SelectedPath + "\\diff.txt");
            if (File.Exists(gitFolder.SelectedPath + "\\diff.txt"))
                File.Delete(gitFolder.SelectedPath + "\\diff.txt");
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
                if(cell.Operation=="Added")
                    listView2.Items[listView2.Items.Count - 1].ForeColor=Color.Green;
                else if (cell.Operation == "Deleted")
                    listView2.Items[listView2.Items.Count - 1].ForeColor = Color.Red;
                else if (cell.Operation == "Changed")
                    listView2.Items[listView2.Items.Count - 1].ForeColor = Color.Orange;
            }
        }

        
        private void listView2_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (select)
            {
                Cursor.Current = Cursors.WaitCursor;
                excelWrapper.OpenFile(gitFolder.SelectedPath + "\\Temp" + extension, cells[e.ItemIndex]);
                select = false;
            }
            else
                excelWrapper.FocusCell(cells[e.ItemIndex]);
        }

        private void listView_MouseMove(object sender, MouseEventArgs e)
        {
            ListView lv = (ListView)sender;
            var hit = lv.HitTest(e.Location);
            if (hit.SubItem != null)
                lv.Cursor = Cursors.Hand;
            else
                lv.Cursor = Cursors.Default;
        }


    }
}