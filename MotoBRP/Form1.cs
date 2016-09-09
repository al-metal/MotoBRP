using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
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
            web.WebRequest webRequest = new web.WebRequest();
            string otv = null;
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

                CookieContainer cookie = Authorizacion();
                string token = tokenReturn(cookie);
                otv = webRequest.getRequest("https://my.tiu.ru/cabinet/product2/index/5187992?status=0", cookie);
                otv = webRequest.getRequest("https://my.tiu.ru/cabinet/product2/create?parent_group=5187992&group=5187992&next=https%3A%2F%2Fmy.tiu.ru%2Fcabinet%2Fproduct2%2Findex%2F5187992%3Fstatus%3D0", cookie);
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

        public CookieContainer Authorizacion()
        {
            CookieContainer cookie = new CookieContainer();

            HttpWebResponse res = null;
            HttpWebRequest req = (HttpWebRequest)System.Net.WebRequest.Create("http://izhevsk.tiu.ru/");
            req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36";
            req.CookieContainer = cookie;
            res = (HttpWebResponse)req.GetResponse();
            StreamReader ressr = new StreamReader(res.GetResponseStream());
            String otv = ressr.ReadToEnd();
            res.GetResponseStream().Close();
            req.GetResponse().Close();
            Cookie ck = res.Cookies["csrf_token"];

            Cookie token = new Cookie("csrf_token", res.Cookies[0].Value);
            string toke = res.Cookies[0].Value;

            req = (HttpWebRequest)HttpWebRequest.Create("https://my.tiu.ru/cabinet/sign-in");
            req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36";
            req.Method = "GET";
            req.ContentType = "application/x-www-form-urlencoded";
            req.CookieContainer = cookie;
            res = (HttpWebResponse)req.GetResponse();
            ressr = new StreamReader(res.GetResponseStream());
            otv = ressr.ReadToEnd();

            req = (HttpWebRequest)HttpWebRequest.Create("https://my.tiu.ru/cabinet/auth/phone_login");
            req.Accept = "application/json, text/javascript, */*; q=0.01";
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36";
            req.Method = "POST";
            req.Referer = "https://my.tiu.ru/cabinet/sign-in";
            req.ContentType = "application/x-www-form-urlencoded";
            req.Headers.Add("X-CSRFToken", toke);
            req.Headers.Add("Origin", "https://my.tiu.ru");
            req.CookieContainer = cookie;
            byte[] ms = Encoding.ASCII.GetBytes("phone_email=moto%40bike18.ru&password=testTIU2016&csrf_token=" + toke + "&_save=YES&");
            req.ContentLength = ms.Length;
            Stream stre = req.GetRequestStream();
            stre.Write(ms, 0, ms.Length);
            stre.Close();
            res = (HttpWebResponse)req.GetResponse();

            return cookie;
        }

        public string tokenReturn(CookieContainer cookie)
        {
            string token = null;

            HttpWebResponse res = null;
            HttpWebRequest req = (HttpWebRequest)System.Net.WebRequest.Create("http://izhevsk.tiu.ru/");
            req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36";
            req.CookieContainer = cookie;
            res = (HttpWebResponse)req.GetResponse();
            StreamReader ressr = new StreamReader(res.GetResponseStream());
            String otv = ressr.ReadToEnd();
            res.GetResponseStream().Close();
            req.GetResponse().Close();
            Cookie ck = res.Cookies["csrf_token"];

            Cookie tokenn = new Cookie("csrf_token", res.Cookies[0].Value);
            token = res.Cookies[0].Value;

            return token;
        }
    }
}
