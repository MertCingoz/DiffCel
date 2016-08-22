using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace EmbeddedExcel
{
    public partial class Form2 : Form
    {
        string path;
        public Form2(string _path)
        {
            path = _path;
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string raw = System.IO.File.ReadAllText(@"raw2.txt");
            raw = raw.Substring(0, raw.IndexOf("----------------- DIFF -------------------"));
            string[] lines = raw.Split('\n');
            excelWrapper.OpenFile(path,lines);
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            excelWrapper.Close();
            excelWrapper.Dispose();
            this.Owner.Show();
            this.Owner.Focus();
        }
    }
}
