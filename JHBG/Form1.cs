using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace JHBG
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            split f1 = new split();
            f1.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            mix f2 = new mix();
            f2.ShowDialog();
        }

        private void 分割ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            split f1 = new split();
            f1.ShowDialog();
        }

        private void 合并ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mix f2 = new mix();
            f2.ShowDialog();
        }

        private void 关于ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 f3 = new AboutBox1();
            f3.ShowDialog();
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
