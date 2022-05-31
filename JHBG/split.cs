using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;

namespace JHBG
{
    public partial class split : Form
    {
        public split()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            label1.Text = null;
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel 工作簿(*.xlsx)|*.xlsx|Excel 97-2003 工作簿(*.xls)|*.xls";
            file.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            file.Multiselect = false;
            if (file.ShowDialog() == DialogResult.Cancel)
                return;
            var path = file.FileName;
            label1.Text = path.ToString();
            string filesuffix = System.IO.Path.GetExtension(path);
            if (string.IsNullOrEmpty(filesuffix))
                return;
            using (DataSet ds = new DataSet())
            {
                string connString = "";
                if (filesuffix == ".xls")
                    connString = "Provider = Microsoft.Jet.OLEDB.4.0;" + "Data Source =" + path + ";" + "; Extended Properties = \"Excel 8.0; HDR = YES; IMEX =1 \"";
                else
                    connString = "Provider = Microsoft.ACE.OLEDB.12.0;" + "Data Source =" + path + ";" + "; Extended Properties = \"Excel 12.0; HDR = YES; IMEX =1 \"";

                //简单读入sheet，后续考虑改为使用Framework提供的对象模型Microsoft.office.Interop.Excel读取第一个表
                string sql_select = "SELECT * FROM[Sheet1$]"; 
                using (OleDbConnection conn = new OleDbConnection(connString))
                using (OleDbDataAdapter cmd = new OleDbDataAdapter(sql_select, conn))
                {
                    conn.Open();
                    cmd.Fill(ds);
                }
                if (ds == null || ds.Tables.Count <= 0) return;
                try
                {
                    dataGridView1.DataSource = ds.Tables[0];

                }
                catch (Exception ex)
                {

                    throw;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //bool title = checkBox1.Checked;
            Workbook bookOriginal = new Workbook();
            bookOriginal.LoadFromFile(label1.Text);
            Worksheet sheet = bookOriginal.Worksheets[0];
            int columnCount = sheet.Columns.Count();
            foreach(CellRange range in sheet.Columns[0])
            {
                Workbook newbook = new Workbook();
                Worksheet newSheet = newbook.Worksheets[0];
                //if (title) { title = false;continue; }
                CellRange titleRange = sheet.Range[1,1,1,columnCount];
                CellRange destRange = newSheet.Range[1, 1, 1, columnCount];
                sheet.Copy(titleRange, destRange,true);
                CellRange sourceRange = sheet.Range[range.Row, 1];
                CellRange destRange2 = newSheet.Range[2, 1];
                sheet.Copy(sourceRange, destRange2, true);
                string saveFileName = range.Columns[0].Text.ToString() + ".xlsx";
                //增加一个文件生成目录
                newbook.SaveToFile(saveFileName, ExcelVersion.Version2013);
                System.Diagnostics.Process.Start(saveFileName);
                //增加进度条
            }
            //增加一个完成提示
        }
    }
}
