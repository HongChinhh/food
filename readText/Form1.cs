using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace readText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            // tao doi tuong mo tep tin
            OpenFileDialog fopen = new OpenFileDialog();
            //chi ra duoi
            fopen.Filter = "(tat ca cac tep)|*.*|(Cac tep excel)|*.xlsx";
            fopen.ShowDialog();
            if(fopen.FileName !="")
            {
                cboPath.Text = fopen.FileName;
                // tao doi tuong
                Excel.Application app = new Excel.Application();
                Excel.Workbook wb = app.Workbooks.Open(fopen.FileName);
                try
                {
                    // mo sheet
                    Excel._Worksheet sheet = wb.Sheets[1];
                    Excel.Range range = sheet.UsedRange;
                    // doc du lieu
                    int rows = range.Rows.Count;
                    int cols = range.Columns.Count;
                    // doc dong tieu de de tao cot cho ListView
                    for(int c =1;c<=cols; c++)
                    {
                        string columname = range.Cells[1, c].Value.ToString();
                        ColumnHeader col = new ColumnHeader();
                        col.Text = columname;
                        col.Width = 120;
                        lstFoodCalculator.Columns.Add(col);
                    }
                    // doc du lieu
                    for(int i=2;i<=rows;i++)
                    {
                        ListViewItem item =  new ListViewItem();
                        for(int j=1;j<cols;j++)
                        {
                            if (j == 1)
                                item.Text = range.Cells[i, j].Value.ToString();
                            else
                                item.SubItems.Add(range.Cells[i, j].Value.ToString());
                        }
                        lstFoodCalculator.Items.Add(item);  
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message, "thong bao", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Ban khong chon tep tin nao", "thong bao", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {

            SaveFileDialog fsave = new SaveFileDialog();
            fsave.Filter = "(tat ca cac tep)|*.*|(Cac tep excel)|*.xlsx";
            fsave.ShowDialog();
            if (fsave.FileName != "")
            {//tao fie exce
                Excel.Application app = new Excel.Application();
                    // tao workbook
                Excel.Workbook wb = app.Workbooks.Add(Type.Missing);
                // tao shet 

            }

                    else
            {
                MessageBox.Show("Ban khong chon tep tin nao", "thong bao", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
