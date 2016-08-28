using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Microsoft.Win32;

namespace EmbeddedExcel
{
    public partial class Form1 : Form
    {
        #region Fields
        public static List<Cell> cells = new List<Cell>();
        private static string[] excelFormats = { "xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xls", ".xlt", ".xml", ".xlam", ".xlw" };
        private FolderBrowserDialog gitFolder = new FolderBrowserDialog();
        private Process cmd = new Process();
        private RegistryKey key = Registry.LocalMachine;
        private string relativePath;
        private string extension;
        private string lastCommit="";
        private int lastCell=-1;
        private string oldPath;
        private string newPath;
        private string selected;
        private TreeNode tempNode;
        private ListViewItem tempItem;
        #endregion Fields

        #region Construction
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            splitContainer4.Panel1Collapsed = true;
            splitContainer5.Panel2Collapsed = true;
            key.CreateSubKey("SOFTWARE");
            key = key.OpenSubKey("SOFTWARE", true);
            key.CreateSubKey("DiffCell");
            key = key.OpenSubKey("DiffCell", true);
            gitFolder.SelectedPath = "Empty";
            foreach (var keyVal in key.GetValueNames())
                if (keyVal == "Path")
                    gitFolder.SelectedPath = key.GetValue(keyVal).ToString();
            
            gitFolder.ShowNewFolderButton = false;
            if (gitFolder.SelectedPath=="Empty")
                gitFolder.ShowDialog();
            ListDirectory(treeView1, gitFolder.SelectedPath);
            treeView1.ExpandAll();
        }
        #endregion Construction

        #region Events
        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected && e.Item.SubItems[0].Text != lastCommit)
            {
                Clean();
                oldLink.LinkVisited = false;
                newLink.LinkVisited = false;
                selected = "";
                splitContainer4.Panel1Collapsed = false;
                splitContainer5.Panel2Collapsed = true;
                listView2.Items.Clear();
                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    string dir = relativePath.Substring(0, relativePath.LastIndexOf("\\") + 1);
                    extension = relativePath.Substring(relativePath.LastIndexOf("."));
                    cmd.Start();
                    cmd.StandardInput.WriteLine("cd " + gitFolder.SelectedPath);
                    if (e.ItemIndex - 1 >= 0)
                        cmd.StandardInput.WriteLine("git diff " + e.Item.SubItems[0].Text + " " + listView1.Items[e.ItemIndex - 1].SubItems[0].Text + " \"" + gitFolder.SelectedPath + "\\" + relativePath + "\" > Temp/diff.txt");
                    else
                        cmd.StandardInput.WriteLine("git diff " + e.Item.SubItems[0].Text + " \"" + gitFolder.SelectedPath + "\\" + relativePath + "\" > Temp/diff.txt");

                    cmd.StandardInput.WriteLine("git cat-file -p " + e.Item.SubItems[0].Text + ":\"" + relativePath.Replace('\\', '/') + "\" > Temp/TempOld" + extension);
                    if (e.ItemIndex - 1 >= 0)
                        cmd.StandardInput.WriteLine("git cat-file -p " + listView1.Items[e.ItemIndex - 1].SubItems[0].Text + ":\"" + relativePath.Replace('\\', '/') + "\" > Temp/TempNew" + extension);

                    cmd.StandardInput.Close();
                    cmd.WaitForExit();

                    listView2.BeginUpdate();
                    GetDiff();

                    oldPath = gitFolder.SelectedPath + "\\Temp\\TempOld" + extension;

                    if (e.ItemIndex == 0)
                        newPath = gitFolder.SelectedPath + "\\" + relativePath;
                    else
                        newPath = gitFolder.SelectedPath + "\\Temp\\TempNew" + extension;
                    lastCommit = e.Item.SubItems[0].Text;
                    listView2.EndUpdate();
                    e.Item.ForeColor = Color.Red;
                    if(tempItem!=null)
                        tempItem.ForeColor = Color.Black;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                tempItem = e.Item;
            }
        }

        private void listView2_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected && e.ItemIndex != lastCell && excelWrapper.Visible)
            {
                excelWrapper.FocusCell(cells[e.ItemIndex]);
                lastCell = e.ItemIndex;
            }
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


        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Nodes.Count == 0 && e.Node.Text != "No file found")
            {
                Clean();
                splitContainer4.Panel1Collapsed = true;
                listView2.Items.Clear();
                relativePath = "";
                TreeNode temp = e.Node;
                while (temp.Parent.Parent != null)
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
                cmd.StandardInput.WriteLine("git log --pretty=format:\"%h|%an|%s|%ci\" \"" + relativePath + "\" > Temp/commits.txt");
                cmd.StandardInput.Close();
                cmd.WaitForExit();

                listView1.Items.Clear();
                string[] lines = System.IO.File.ReadAllLines(gitFolder.SelectedPath + "\\Temp\\commits.txt");

                string author = "";
                string description = "current";
                string date = "";
                for (int i = 0; i < lines.Length; i++)
                {
                    string[] item = lines[i].Split('|');
                    listView1.Items.Add(item[0]);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(author);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(description);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(date);
                    author = item[1];
                    description = item[2];
                    date = item[3].Substring(0, item[3].Length - 5);
                }


                if (tempNode != null)
                    tempNode.ForeColor = Color.Black;
                tempNode = e.Node;
                tempNode.ForeColor = Color.Red;
            }
            else if (e.Node.Parent == null)
            {
                string temp = gitFolder.SelectedPath;
                gitFolder.ShowDialog();
                if (temp != gitFolder.SelectedPath)
                {
                    ListDirectory(treeView1, gitFolder.SelectedPath);
                    listView1.Items.Clear();
                    listView2.Items.Clear();
                    excelWrapper.Visible = false;
                    lastCommit = "";
                }
            }
        }

        private void openWithLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Clean();
            Process.Start(selected);
            selected = "";
            splitContainer5.Panel2Collapsed = true;
            oldLink.LinkVisited = false;
            newLink.LinkVisited = false;
            openWithLink.LinkVisited = true;
        }

        private void oldLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (selected != oldPath)
            {
                Clean();
                excelWrapper.OpenFile(oldPath, false);
                splitContainer5.Panel2Collapsed = false;
                selected = oldPath;
                oldLink.LinkVisited = true;
                newLink.LinkVisited = false;
                openWithLink.LinkVisited = false;
            }
        }

        private void newLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (selected != newPath)
            {
                Clean();
                excelWrapper.OpenFile(newPath, false);
                splitContainer5.Panel2Collapsed = false;
                selected = newPath;
                oldLink.LinkVisited = false;
                newLink.LinkVisited = true;
                openWithLink.LinkVisited = false;
            }
        }

        #endregion Events

        #region Methods
        private void ListDirectory(TreeView treeView, string path)
        {
            tempNode = null;
            key.SetValue("Path", path);
            treeView.Nodes.Clear();
            var rootDirectoryInfo = new DirectoryInfo(path);
            TreeNode root = new TreeNode("Select Git Folder");
            TreeNode repo = null;
            try
            {
                repo = CreateDirectoryNode(rootDirectoryInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            if (repo != null) root.Nodes.Add(repo);
            treeView.Nodes.Add(root);
            if (treeView.Nodes[0].Nodes.Count == 0) treeView.Nodes[0].Nodes.Add("No file found");
            treeView.SelectedNode = treeView.Nodes[0].Nodes[0];
        }

        private static TreeNode CreateDirectoryNode(DirectoryInfo directoryInfo)
        {
            var directoryNode = new TreeNode(directoryInfo.Name);
            foreach (var directory in directoryInfo.GetDirectories())
                if (CreateDirectoryNode(directory) != null)
                    directoryNode.Nodes.Add(CreateDirectoryNode(directory));
            foreach (var file in directoryInfo.GetFiles())
                foreach (var item in excelFormats)
                    if (file.Name.Contains(item) && !file.Name.Contains("Temp"))
                        directoryNode.Nodes.Add(new TreeNode(file.Name));
            if (directoryNode.Nodes.Count == 0)
                return null;
            return directoryNode;
        }

        private void GetDiff()
        {
            try
            {
                cells.Clear();
                string raw = System.IO.File.ReadAllText(gitFolder.SelectedPath + "\\Temp\\diff.txt");
                if (raw.Length == 0) return;
                raw = raw.Substring(0, raw.IndexOf("----------------- DIFF -------------------") - 1);
                string[] lines = raw.Split('\n');
                foreach (var line in lines)
                {
                    Cell cell = new Cell();
                    cell.Sheet = line.Substring(18, line.IndexOf("!", 18) - 18);
                    cell.Adress = line.Substring(line.IndexOf("!", 18) + 1, line.IndexOf(" ", 18) - line.IndexOf("!", 18) - 1);
                    if (line.Substring(14, 3) == "   ")
                    {
                        cell.OldValue = line.Substring(line.IndexOf("=>") + 4, line.IndexOf("' v/s '") - line.IndexOf("=>") - 4);
                        cell.NewValue = line.Substring(line.IndexOf("' v/s '") + 7, line.Length - line.IndexOf("' v/s '") - 9);
                        cell.Operation = "Change";
                    }
                    else if (line.Substring(14, 3) == "WB1")
                    {
                        cell.OldValue = line.Substring(line.IndexOf("=>") + 4, line.Length - line.IndexOf("=>") - 6);
                        cell.Operation = "Delete";
                    }
                    else if (line.Substring(14, 3) == "WB2")
                    {
                        cell.NewValue = line.Substring(line.IndexOf("=>") + 4, line.Length - line.IndexOf("=>") - 6);
                        cell.Operation = "Add";
                    }
                    cells.Add(cell);
                    listView2.Items.Add(cell.Operation);
                    listView2.Items[listView2.Items.Count - 1].SubItems.Add(cell.Sheet);
                    listView2.Items[listView2.Items.Count - 1].SubItems.Add(cell.Adress);
                    listView2.Items[listView2.Items.Count - 1].SubItems.Add(cell.OldValue);
                    listView2.Items[listView2.Items.Count - 1].SubItems.Add(cell.NewValue);
                    if (cell.Operation == "Add")
                        listView2.Items[listView2.Items.Count - 1].ForeColor = Color.Green;
                    else if (cell.Operation == "Delete")
                        listView2.Items[listView2.Items.Count - 1].ForeColor = Color.Red;
                    else if (cell.Operation == "Change")
                        listView2.Items[listView2.Items.Count - 1].ForeColor = Color.Orange;
                }
            }
            catch
            { }
        }

        private void Clean()
        {
            excelWrapper.Dispose();
            excelWrapper = new EmbeddedExcel.ExcelWrapper();
            excelWrapper.Dock = System.Windows.Forms.DockStyle.Fill;
            excelWrapper.Location = new System.Drawing.Point(0, 0);
            excelWrapper.Margin = new System.Windows.Forms.Padding(5);
            excelWrapper.Name = "excelWrapper";
            excelWrapper.Size = new System.Drawing.Size(697, 628);
            excelWrapper.TabIndex = 11;
            excelWrapper.Visible = false;
            splitContainer4.Panel2.Controls.Add(excelWrapper);
        }
        #endregion Methods


    }
}