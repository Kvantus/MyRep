using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using mshtml;
using IWshRuntimeLibrary;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace RazvLogins
{


    public partial class Form1 : Form
    {
        static public WshShell cmd = null;
        static SHDocVw.InternetExplorer IE = null;

        public Form1()
        {
            //DateTime myDate = DateTime.ParseExact("2009-05-08 14:40:52,531", "yyyy-MM-dd HH:mm:ss,fff",
            //                           System.Globalization.CultureInfo.InvariantCulture);
            //MessageBox.Show(myDate.Day.ToString() + " " + myDate.Month.ToString());
            //MessageBox.Show(DateTime.Now + " - " + myDate);

                InitializeComponent();
                if (Environment.UserName == "viktor_k")
                {
                    BTest.Visible = true;
                }

        }
        
        public ManagersSups managers = new ManagersSups();

        static public void GOLogin1(string email, string pass)
        {
            Process chromik = Process.Start("chrome.exe", "http://supplier.autoeuro.ru/login");


            //MessageBox.Show(chromik.GetType().ToString());
        }

        static public void GOLogin (string email, string pass)
        {
            Process chromik;
            try
            {
                if (Environment.UserName.ToLower() == "brusakova_n")
                {
                    chromik = Process.Start(@"C:\Documents and Settings\Brusakova_n\Local Settings\Application Data\Google\Chrome\Application\chrome.exe", "http://supplier.autoeuro.ru/login");
                }
                else if (Environment.UserName.ToLower() == "Lipatova_E".ToLower())
                {
                    chromik = Process.Start(@"C:\Documents and Settings\Lipatova_E\Local Settings\Application Data\Google\Chrome\Application\chrome.exe", "http://supplier.autoeuro.ru/login");
                }
                else
                {
                    chromik = Process.Start("chrome.exe", "http://supplier.autoeuro.ru/login");

                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Не получилось запустить Хром, пичалька :(\n" + e.Message);
                return;
            }
            //chromik.WaitForInputIdle(2500);
            Application.DoEvents();
            if (Environment.UserName.ToLower() == "embakh_a")
            {
                Thread.Sleep(1000);
            }
            else
            {
                Thread.Sleep(1800);
            }
            
            Clipboard.SetText(email);
            SendKeys.SendWait("+{INSERT}");
            Application.DoEvents();
            SendKeys.SendWait("{TAB}");
            Thread.Sleep(50);
            Application.DoEvents();
            Clipboard.SetText(pass);
            
            SendKeys.SendWait("+{INSERT}");
            Application.DoEvents();
            SendKeys.SendWait("{ENTER}");
            //SendKeys.SendWait("+{HOME}");
            //Application.DoEvents();
            //System.Threading.Thread.Sleep(100);
            //SendKeys.SendWait(pass);

            

            //второй заход для особо тупых случаев
            //SendKeys.SendWait("{TAB}");
            //Application.DoEvents();
            //SendKeys.SendWait("+{INSERT}");
            //Application.DoEvents();
            //SendKeys.SendWait("{ENTER}");

        }

        static public void GOLogin2(string email, string pass)
        {
            string login = email;
            string password = pass;

            IE = new SHDocVw.InternetExplorer();
            IE.Visible = true;

            IE.Navigate("http://supplier.autoeuro.ru/login");

            while (IE.ReadyState.ToString() != "READYSTATE_COMPLETE")
            {
                Application.DoEvents();
            }

            HTMLDocument doc = (HTMLDocument)IE.Document;
            IHTMLElement pole1 = doc.getElementById("email");
            pole1.innerText = login;
            IHTMLElement pole2 = doc.getElementById("password");
            pole2.innerText = password;
            HTMLButtonElement button = (HTMLButtonElement)doc.getElementsByTagName("button").item(0);
            button.click();
        }


        private void BTest_Click(object sender, EventArgs e)
        {
            ManagersSups managers = new ManagersSups();
            managers.GetSups2();


            //var chromik = Process.Start("chrome.exe", "http://supplier.autoeuro.ru/login");
            
            //string login = "volvo-auto@autoeuro.ru";
            //string password = "volvo-auto@autoeuro.ru";

            //IE = new SHDocVw.InternetExplorer();
            //IE.Visible = true;

            //IE.Navigate("http://supplier.autoeuro.ru/login");
            ////System.Threading.Thread.Sleep(3000);

            //while (IE.ReadyState.ToString() != "READYSTATE_COMPLETE")
            //{
            //    Application.DoEvents();
            //}

            ////Type IEType = Type.GetTypeFromProgID("InternetExplorer.Application");
            ////dynamic IE = Activator.CreateInstance(IEType);

            ////IE.Navigate("http://supplier.autoeuro.ru/login");
            ////IE.Visible = true;

            ////HTMLDocument dic = IE.


            //HTMLDocument doc = IE.Document;
            //IHTMLElement pole1 = doc.getElementById("email");
            //pole1.innerText = login;
            //IHTMLElement pole2 = doc.getElementById("password");
            //pole2.innerText = password;
            ////var button = doc.getElementsByTagName("input");
            //HTMLButtonElement button = doc.getElementsByTagName("button").item(0);
            //button.click();

            ////MessageBox.Show(button.GetType().ToString());


        }

        private void BEnd_Click(object sender, EventArgs e)
        {
            //IE.Quit();
            IE = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            this.Close();
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            managers.GetSups2();

            CreateButtons(managers.DiBrusakova, 0);
            CreateButtons(managers.DiLA, 1);
            CreateButtons(managers.DiLK, 2);
            CreateButtons(managers.DiRizhkova, 3);
            CreateButtons(managers.DiEmbah, 4);

        }

        private void CreateButtons(SortedDictionary<string, string> manager , int tabPageNumber)
        {
            int counter = 0;
            int locX = 15;
            int locY = 15;
            foreach (var item in manager)
            {
                Button a = new Button();
                a.Size = new Size(300, 30);
                a.Text = item.Key;
                a.Location = new Point(locX, locY);

                a.Click += new EventHandler(Smart_Click);
                SmartMultiPage.TabPages[tabPageNumber].Controls.Add(a);
                counter++;
                if (counter % 12 == 0)
                {
                    locX += 310;
                    locY = 15;
                }
                else
                {
                    locY += 35;
                }

            }
        }


        private void Smart_Click(object sender, EventArgs e)
        {

            Button button = (Button)sender;
            string supplier = button.Text;
            string login = null;
            if (managers.DiBrusakova.ContainsKey(supplier))
            {
                login = managers.DiBrusakova[supplier];
            }
            else if (managers.DiLA.ContainsKey(supplier))
            {
                login = managers.DiLA[supplier];
            }
            else if (managers.DiLK.ContainsKey(supplier))
            {
                login = managers.DiLK[supplier];
            }
            else if (managers.DiRizhkova.ContainsKey(supplier))
            {
                login = managers.DiRizhkova[supplier];
            }
            else if (managers.DiEmbah.ContainsKey(supplier))
            {
                login = managers.DiEmbah[supplier];
            }
            else
            {
                return;
            }

            GOLogin(login, login);
        }

    }
}
