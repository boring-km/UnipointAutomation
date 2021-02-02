using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Interop.mshtml;
using SHDocVw;

namespace UnipointAutomation
{
    public partial class Form1 : Form
    {
        string id = "";
        string pw = "";
        public Form1()
        {
            InitializeComponent();
        }

        // TODO: 아이디/비밀번호 불러오기
        private void Form1_Load(object sender, EventArgs e)
        {

            StreamReader reader = null;
            try
            {
                reader = new StreamReader(@"c:/attend_auto/idpw.txt", System.Text.Encoding.GetEncoding("utf-8"), true);
                id = reader.ReadLine();
                pw = reader.ReadLine();
                textBox1.Text = id;
                textBox2.Text = pw;
                button2.PerformClick();
            } catch
            {
                label1.Text = "저장된 아이디 비밀번호 없음";
            } finally
            {
                if(reader != null)
                    reader.Close();
            }
        }

        // TODO: 아이디/비밀번호 저장
        private void button1_Click(object sender, EventArgs e)
        {
            string inputId = textBox1.Text;
            string inputPw = textBox2.Text;

            string folderPath = "C:/attend_auto";
            DirectoryInfo di = new DirectoryInfo(folderPath);
            if(di.Exists == false)
            {
                di.Create();
            }
            StreamWriter writer = null;
            try
            {
                writer = new StreamWriter(@"c:/attend_auto/idpw.txt", true, System.Text.Encoding.GetEncoding("utf-8"));
                writer.WriteLine(inputId);
                writer.WriteLine(inputPw);
                label2.Text = "C:/attend_auto/idpw.txt 에 저장됨";
            } catch
            {
                label1.Text = "파일 저장 에러";
            } finally
            {
                writer.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(id == "" || pw == "")
            {
                return;
            }
            label1.Text = "인터넷 여는 중...";
            var ie = new InternetExplorer();
            var webBrowser = (IWebBrowserApp)ie;
            webBrowser.Visible = true;
            webBrowser.Navigate("http://gw.unipoint.co.kr");
            label1.Text = "아이디 비밀번호 입력 대기중...";
            Thread.Sleep(3000);

            IHTMLDocument3 doc = (IHTMLDocument3)webBrowser.Document;
            IHTMLElement idElement = doc.getElementById("UserName");
            IHTMLElement pwElement = doc.getElementById("Password");
            idElement.setAttribute("value", id);
            pwElement.setAttribute("value", pw);
            doc.getElementsByTagName("a").item(0).click();
            label1.Text = "로그인 중...";
            Thread.Sleep(1500);

            IHTMLDocument7 doc7 = (IHTMLDocument7)webBrowser.Document;
            IHTMLElement timeElement = doc7.getElementsByClassName("time_num").item(0);
            String checkTime = doc7.getElementsByClassName("time_num").item(0).getAttribute("innerText");
            if (checkTime.Length > 4)
            {
                label1.Text = "퇴근 처리중...";
                // doc.getElementById("btnAttOut").click();
                Thread.Sleep(1000);
                timeElement = doc7.getElementsByClassName("time_num").item(1);
                String endTime = timeElement.getAttribute("innerText");
                if(endTime != null)
                {
                    label1.Text = "퇴근 처리 완료 및 프로그램 종료...";
                    this.Close();
                } else
                {
                    label1.Text = "퇴근 처리 실패";
                    Thread.Sleep(1000);
                }
            }
            else
            {
                label1.Text = "출근 처리중...";
                // doc.getElementById("btnAttOut").click();
                Thread.Sleep(1000);
                timeElement = doc7.getElementsByClassName("time_num").item(0);
                String startTime = timeElement.getAttribute("innerText");
                if(startTime != null)
                {
                    label1.Text = "출근 처리 완료 및 프로그램 종료...";
                    this.Close();
                } else
                {
                    label1.Text = "출근 처리 실패";
                    Thread.Sleep(1000);
                }
            }
        }
    }
}
