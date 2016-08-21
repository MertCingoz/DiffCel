using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace EmbeddedExcel
{
	public partial class Form1 : Form
	{
        Process cmd = new Process();
        string path;

		public Form1() {
			InitializeComponent();
		}

        private void Form1_Load(object sender, EventArgs e)
        {
            ListDirectory(treeView1, AppDomain.CurrentDomain.BaseDirectory);
        }

        private void ListDirectory(TreeView treeView, string path)
        {
            treeView.Nodes.Clear();
            var rootDirectoryInfo = new DirectoryInfo(path);
            treeView.Nodes.Add(CreateDirectoryNode(rootDirectoryInfo));
        }

        private static TreeNode CreateDirectoryNode(DirectoryInfo directoryInfo)
        {
            var directoryNode = new TreeNode(directoryInfo.Name);
            foreach (var directory in directoryInfo.GetDirectories())
                directoryNode.Nodes.Add(CreateDirectoryNode(directory));
            foreach (var file in directoryInfo.GetFiles())
                if(file.Name.Contains(".xls"))
                    directoryNode.Nodes.Add(new TreeNode(file.Name));
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
                    path = "\\" + temp.Text+path;
                    temp = temp.Parent;
                }
                path=path.Substring(1,path.Length-1);
                
                cmd.StartInfo.FileName = "cmd.exe";
                cmd.StartInfo.RedirectStandardInput = true;
                cmd.StartInfo.RedirectStandardOutput = true;
                cmd.StartInfo.CreateNoWindow = true;
                cmd.StartInfo.UseShellExecute = false;
                cmd.Start();
                cmd.StandardInput.WriteLine("git log --pretty=oneline --abbrev-commit path >raw.txt");
                cmd.StandardInput.Close();
                cmd.WaitForExit();

                string[] lines = System.IO.File.ReadAllLines(@"raw.txt");
                foreach (string line in lines)
                {
                    int index = line.IndexOf(' ');
                    listView1.Items.Add(line.Substring(0, index));
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(line.Substring(index + 1, line.Length - index - 1));
                }

                treeView1.Visible = false;
                listView1.Visible = true;
            }
        }

        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected)
            {
                cmd.Start();
                cmd.StandardInput.WriteLine("git diff " + e.Item.SubItems[0].Text + " TestWorkbook.xlsx >raw2.txt");
                cmd.StandardInput.Close();
                cmd.WaitForExit();

                textBox1.Text = System.IO.File.ReadAllText(@"raw2.txt");
                excelWrapper1.OpenFile(AppDomain.CurrentDomain.BaseDirectory + path);
            }
        }



	}
}