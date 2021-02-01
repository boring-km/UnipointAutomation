using System;
using System.Threading;
using System.Windows.Forms;
using mshtml;
using SHDocVw;

namespace UnipointAutomation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var ie = new InternetExplorer();
            var webBrowser = (IWebBrowserApp)ie;
            webBrowser.Visible = true;
            webBrowser.Navigate("http://gw.unipoint.co.kr");
            Thread.Sleep(3000);
            IHTMLDocument3 doc = (IHTMLDocument3)webBrowser.Document;
            IHTMLElement id = doc.getElementById("UserName");
            IHTMLElement pw = doc.getElementById("Password");
            id.setAttribute("value", "아이디");
            pw.setAttribute("value", "비밀번호");
            IHTMLElementCollection tagList = doc.getElementsByTagName("a");
            foreach (IHTMLElement tag in tagList)
            {
                tag.click();
                break;
            }
        }

    }
}
