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
            FileInfo file = new FileInfo("Копия остатки для bike18.xlsx");
            ExcelPackage p = new ExcelPackage(file);
            ExcelWorksheet w = p.Workbook.Worksheets[1];
            int q = w.Dimension.Rows;
            for (int i = 2; q > i; i++)
            {
                string name = w.Cells[i, 2].Value.ToString();
                string article = w.Cells[i, 1].Value.ToString().Trim();
                double price = (double)w.Cells[i, 4].Value;
                int priceInSite = formulaPrice(price);
            }
        }

        private int formulaPrice(double price)
        {
            int priceInSite = 0;
            if (price <= 15)
                price = price * 2.7;
            else
            if (price <= 199)
                price = price * 2.5;
            else
            if (price <= 2000)
                price = price * 1.7;
            else
            if (price <= 7999)
                price = price * 1.3;
            else
            if (price > 8000)
                price = price * 1.2;

            priceInSite = Convert.ToInt32(price);
            priceInSite = (priceInSite + 5) / 10 * 10;

            return priceInSite;
        }
    }
}
