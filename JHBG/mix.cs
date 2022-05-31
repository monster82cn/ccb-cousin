using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;

namespace JHBG
{
    public partial class mix : Form
    {
        public mix()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            label1.Text = null;
            OpenFileDialog files = new OpenFileDialog();
            files.Filter = "Excel 工作簿(*.xlsx)|*.xlsx|Excel 97-2003 工作簿(*.xls)|*.xls";
            files.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            files.Multiselect = true;
            if (files.ShowDialog() == DialogResult.Cancel)
                return;
            int newRow = 2;
            Workbook mixFile = new Workbook();
            Worksheet mixSheet = mixFile.Worksheets[0];
            foreach(var file in files.FileNames)
            {
                Workbook eachFile = new Workbook();
                eachFile.LoadFromFile(file.ToString());
                Worksheet eachSheet = eachFile.Worksheets[0];
                int columnCount = eachSheet.Columns.Count();
                for(int i=1;i<=eachSheet.Rows.Count();i++)
                {
                    if (i == 1) {
                        CellRange titleRange = eachSheet.Range[1, 1, 1, columnCount];
                        CellRange desttitleRange = mixSheet.Range[1, 1, 1, columnCount];
                        mixSheet.Copy(titleRange, desttitleRange, true);
                        continue;
                    }
                    CellRange eachRange = eachSheet.Range[i, 1, i, columnCount];
                    CellRange destRange = mixSheet.Range[newRow, 1, newRow, columnCount];
                    mixSheet.Copy(eachRange, destRange, true);
                    newRow++;
                }
                label1.Text = label1.Text + file.ToString() + ";";
                mixFile.SaveToFile("mix.xlsx", ExcelVersion.Version2013);
                System.Diagnostics.Process.Start("mix.xlsx");

            }
        }
    }
}
