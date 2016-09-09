using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MotoBRP
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnActualPrice_Click(object sender, EventArgs e)
        {
            FileInfo file = new FileInfo("Запчасти.xlsx");
            ExcelPackage p = new ExcelPackage(file);
            ExcelWorksheet w = p.Workbook.Worksheets[1];
            int q = w.Dimension.Rows;
            for (int i = 14; q > i; i++)
            {

            }
        }
    }
}
