﻿using System;
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
using Microsoft.Win32;

namespace EmbeddedExcel
{
    public partial class Form1 : Form
    {
        public static List<Cell> cells = new List<Cell>();
        private static string[] excelFormats = { "xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xls", ".xlt", ".xml", ".xlam", ".xlw" };
        private FolderBrowserDialog gitFolder = new FolderBrowserDialog();
        private Process cmd = new Process();
        private RegistryKey key = Registry.LocalMachine;
        private string relativePath;
        private string extension;
        private string lastCommit="";
        private int lastCell=-1;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
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

        private void ListDirectory(TreeView treeView, string path)
        {
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
                if (excelFormats.Any(file.Name.Contains) && !file.Name.Contains("Temp"))
                    directoryNode.Nodes.Add(new TreeNode(file.Name));
            if (directoryNode.Nodes.Count == 0)
                return null;
            return directoryNode;
        }

        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected && e.Item.SubItems[0].Text!=lastCommit)
            {
                listView2.Items.Clear();
                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    excelWrapperOld.Dispose();
                    excelWrapperOld = new EmbeddedExcel.ExcelWrapper();
                    excelWrapperOld.Dock = System.Windows.Forms.DockStyle.Fill;
                    excelWrapperOld.Location = new System.Drawing.Point(0, 0);
                    excelWrapperOld.Margin = new System.Windows.Forms.Padding(5);
                    excelWrapperOld.Name = "excelWrapperOld";
                    excelWrapperOld.Size = new System.Drawing.Size(318, 626);
                    excelWrapperOld.TabIndex = 9;
                    excelWrapperOld.Visible = false;

                    excelWrapperNew.Dispose();
                    excelWrapperNew = new EmbeddedExcel.ExcelWrapper();
                    excelWrapperNew.Dock = System.Windows.Forms.DockStyle.Fill;
                    excelWrapperNew.Location = new System.Drawing.Point(0, 0);
                    excelWrapperNew.Margin = new System.Windows.Forms.Padding(5);
                    excelWrapperNew.Name = "excelWrapperNew";
                    excelWrapperNew.Size = new System.Drawing.Size(319, 626);
                    excelWrapperNew.TabIndex = 10;
                    excelWrapperNew.Visible = false;

                    splitContainer4.Panel1.Controls.Add(excelWrapperOld);
                    splitContainer4.Panel2.Controls.Add(excelWrapperNew);

                    string[] files = System.IO.Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "TempOld*", System.IO.SearchOption.TopDirectoryOnly);
                    foreach (var file in files)
                        File.Delete(file);

                    string dir = relativePath.Substring(0, relativePath.LastIndexOf("\\") + 1);
                    extension = relativePath.Substring(relativePath.LastIndexOf("."));
                    cmd.Start();
                    cmd.StandardInput.WriteLine("cd " + gitFolder.SelectedPath);
                    cmd.StandardInput.WriteLine("git diff " + e.Item.SubItems[0].Text + " \"" + gitFolder.SelectedPath + "\\" + relativePath + "\" > Temp/diff.txt");
                    cmd.StandardInput.WriteLine("git cat-file -p " + e.Item.SubItems[0].Text + ":\"" + relativePath.Replace('\\', '/') + "\" > Temp/TempOld" + extension);
                    if (e.ItemIndex-1>0)
                        cmd.StandardInput.WriteLine("git cat-file -p " + listView1.Items[e.ItemIndex -1].SubItems[0].Text + ":\"" + relativePath.Replace('\\', '/') + "\" > Temp/TempNew" + extension);

                    cmd.StandardInput.Close();
                    cmd.WaitForExit();

                    listView2.BeginUpdate();
                    GetDiff();
                    new Thread(() =>
                    {
                        excelWrapperOld.OpenFile(gitFolder.SelectedPath + "\\Temp\\TempOld" + extension);
                        if (e.ItemIndex - 1 > 0)
                            excelWrapperNew.OpenFile(gitFolder.SelectedPath + "\\Temp\\TempNew" + extension);

                    }).Start();
                    lastCommit = e.Item.SubItems[0].Text;
                    listView2.EndUpdate();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void GetDiff()
        {
            try
            {
                cells.Clear();
                string raw = System.IO.File.ReadAllText(gitFolder.SelectedPath + "\\Temp\\diff.txt");
                if (File.Exists(gitFolder.SelectedPath + "\\Temp\\diff.txt"))
                    File.Delete(gitFolder.SelectedPath + "\\Temp\\diff.txt");
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
                    if (cell.Operation == "Added")
                        listView2.Items[listView2.Items.Count - 1].ForeColor = Color.Green;
                    else if (cell.Operation == "Deleted")
                        listView2.Items[listView2.Items.Count - 1].ForeColor = Color.Red;
                    else if (cell.Operation == "Changed")
                        listView2.Items[listView2.Items.Count - 1].ForeColor = Color.Orange;
                }
            }
            catch
            {  }
        }

        private void listView2_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected && e.ItemIndex != lastCell)
            {
                excelWrapperOld.FocusCell(cells[e.ItemIndex]);
                excelWrapperNew.FocusCell(cells[e.ItemIndex]);
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

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Nodes.Count == 0 && e.Node.Text != "No file found")
            {
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
                cmd.StandardInput.WriteLine("git log --pretty=format:\"%h|%an|%s\" \"" + relativePath + "\" > Temp/commits.txt");
                cmd.StandardInput.Close();
                cmd.WaitForExit();

                listView1.Items.Clear();
                string[] lines = System.IO.File.ReadAllLines(gitFolder.SelectedPath + "\\Temp\\commits.txt");
                if (File.Exists(gitFolder.SelectedPath + "\\Temp\\commits.txt"))
                    File.Delete(gitFolder.SelectedPath + "\\Temp\\commits.txt");

                for (int i = 0; i < lines.Length; i++)
                {
                    string[] item = lines[i].Split('|');
                    listView1.Items.Add(item[0]);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(item[1]);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(item[2]);
                }
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
                    excelWrapperOld.Visible = false;
                    excelWrapperNew.Visible = false;
                    lastCommit = "";
                }
            }
        }
    }
}