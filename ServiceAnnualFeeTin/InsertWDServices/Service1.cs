﻿using Microsoft.Reporting.WebForms;
using MimeTypes;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using static ConsoleApplication2.ResponsWhatsapp;

namespace InsertWDServices
{
    public partial class Service1 : ServiceBase
    {
        public static string fileId = "";
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("started at " + DateTime.Now);
            string file = AppDomain.CurrentDomain.BaseDirectory + "\\Report\\numberregistered.txt";
            string[] text = File.ReadAllLines(file);
            List<string> number = new List<string>();
            foreach (string item in text)
            {
                string[] numberdata = item.Split(' ');
                number.Add(numberdata[0]);
            }
            foreach (var item in number)
            {
                SendMessage(item, "System reminder annual fee timah start TimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
            }

            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 3600000;  //1 jam
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();
        }
        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            try
            {
                List<int> messageNumber = new List<int>();

                //masukin last message numbernya
                var day = DateTime.Now.ToString("dd");
                var month = DateTime.Now.ToString("MM");
                var jam = DateTime.Now.ToString("HH");
                var dr1 = new ConsoleApplication2.BankTrasferTableAdapters.mst_srtifikat_kepesertaanTableAdapter();
                var dt1 = dr1.GetData();
                foreach (var item in dt1)
                {
                    var dayDb = item.tanggal.ToString("dd");
                    var monthDb = item.tanggal.ToString("MM");
                    if (day == dayDb && month == monthDb)
                    {
                        if (jam == "08")
                        {
                            getInvoice(item.Nama, (item.CMID).ToString(), dayDb, monthDb);
                        }
                        string file = AppDomain.CurrentDomain.BaseDirectory + "\\Report\\numberregistered.txt";
                        string[] text = File.ReadAllLines(file);
                        List<string> number = new List<string>();
                        foreach (string item_number in text)
                        {
                            string[] numberdata = item_number.Split(' ');
                            number.Add(numberdata[0]);
                        }
                        foreach (var item_number in number)
                        {
                            SendMessage(item_number, "Reminder annual fee"+item.Nama+", Invoice akan di generate jam 08.00 di telegram\nTimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
                        }
                    }
                }
            }
            catch (Exception x)
            {
                WriteToFile(x.Message);
            }

        }

        protected override void OnStop()
        {
            string file = AppDomain.CurrentDomain.BaseDirectory + "\\Report\\numberregistered.txt";
            string[] text = File.ReadAllLines(file);
            List<string> number = new List<string>();
            foreach (string item in text)
            {
                string[] numberdata = item.Split(' ');
                number.Add(numberdata[0]);
            }
            foreach (var item in number)
            {
                SendMessage(item, "System reminder annual fee timah stopped TimeStamp : " + DateTime.Now.ToString("hh:mm:ss"));
            }
            WriteToFile("stoped at " + DateTime.Now);
        }
        public static void getInvoice(string namefile, string cmid, string day, string month)
        {
            var startdate = DateTime.Now.AddYears(-1).ToString("yyyy") + "-" + month + "-" + day;
            var enddate = DateTime.Now.ToString("yyyy") + "-" + month + "-" + day;
            if (Convert.ToDateTime(enddate).DayOfWeek == DayOfWeek.Saturday)
            {
                enddate = Convert.ToDateTime(enddate).AddDays(1).ToString("yyyy-MM-dd");
                if (Convert.ToDateTime(enddate).DayOfWeek == DayOfWeek.Sunday)
                {
                    enddate = Convert.ToDateTime(enddate).AddDays(1).ToString("yyyy-MM-dd");
                }
            }
            string fileSeq = AppDomain.CurrentDomain.BaseDirectory + "\\Report\\seq.txt";
            string[] byteSeq = System.IO.File.ReadAllLines(fileSeq);
            int seq = Convert.ToInt32(byteSeq[0]);

            List<String> filePathAll = new List<string>();

            //INVOICE SELLER
            string inv = "00" + seq + "/KBI/TIMAH/" + month + "/" + DateTime.Now.ToString("yyyy");

            string filePath = getReportSSRSWord("InvoiceAnnualFeeSeller", " &invnumber=" + inv + "&tgl=" + startdate + "&CMID=" + cmid + "&tglend=" + enddate, "Invoice Annual Fee Seller " + namefile, "TIN_EOD_Report");
            sendFileTelegram("-1001649045625", filePath);

            //INVOICE BUYER
            seq = seq + 1;
            inv = "00" + seq + "/KBI/TIMAH/" + month + "/" + DateTime.Now.ToString("yyyy");
            filePath = getReportSSRSWord("InvoiceAnnualFeeBuyer", " &invnumber=" + inv + "&tgl=" + startdate + "&CMID=" + cmid + "&tglend=" + enddate, "Invoice Annual Fee Buyer " + namefile, "TIN_EOD_Report");
            sendFileTelegram("-1001649045625", filePath);

            //PROFORMA SELLER
            seq = seq + 1;
            inv = "00" + seq + "/KBI/PROFORMATIMAH/" + month + "/" + DateTime.Now.ToString("yyyy");
            filePath = getReportSSRSWord("ProformaAnnualFeeSeller", " &invnumber=" + inv + "&tgl=" + startdate + "&CMID=" + cmid + "&tglend=" + enddate, "Proforma Seller " + namefile, "TIN_EOD_Report");
            sendFileTelegram("-1001649045625", filePath);

            //PROFORMA BUYER
            seq = seq + 1;
            inv = "00" + seq + "/KBI/PROFORMATIMAH/" + month + "/" + DateTime.Now.ToString("yyyy"); ;

            filePath = getReportSSRSWord("ProformaAnnualFeeBuyer", " &invnumber=" + inv + "&tgl=" + startdate + "&CMID=" + cmid + "&tglend=" + enddate, "Proforma Buyer " + namefile, "TIN_EOD_Report");
            sendFileTelegram("-1001649045625", filePath);

            string text2 = seq.ToString();
            System.IO.File.WriteAllText(fileSeq, text2);
        }
        public static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!System.IO.File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = System.IO.File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = System.IO.File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
        private static void sendFileTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls12;
            var client = new RestClient("https://api.telegram.org/bot5278461864:AAGwGEV3aJx8ZFKFfRaFB6TMc4mOem-uoc8/sendDocument");
            client.Timeout = -1;
            var requestWa = new RestRequest(Method.POST);
            requestWa.AddHeader("Content-Type", "multipart/form-data");
            requestWa.AddParameter("chat_id", chatId);
            requestWa.AddFile("document", body);
            IRestResponse responseWa = client.Execute(requestWa);
            Console.WriteLine(responseWa.Content);
        }
        private static void SendMessage(string chatId, string body)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/sendMessage?token=jkdjtwjkwq2gfkac");
            client.Timeout = -1;
            var requestWa = new RestRequest(Method.POST);
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("phone", chatId);
            requestWa.AddParameter("body", body);
            IRestResponse responseWa = client.Execute(requestWa);
        }
        public static string getReportSSRSWord(string reportname, string param, string filename, string pathreport)
        {
            try
            {
                string url = "http://10.15.2.60/ReportServerEOD?/" + pathreport + "/" + reportname + "&rs:Command=Render&rs:Format=WORD&rc:OutputFormat=DOCX" + param;

                System.Net.HttpWebRequest Req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                Req.Credentials = new NetworkCredential("Administrator", "Jakarta01");
                Req.Method = "GET";

                string path = AppDomain.CurrentDomain.BaseDirectory + "report\\" + filename + ".doc";

                System.Net.WebResponse objResponse = Req.GetResponse();
                System.IO.FileStream fs = new System.IO.FileStream(path, System.IO.FileMode.Create);
                System.IO.Stream stream = objResponse.GetResponseStream();

                byte[] buf = new byte[1024];
                int len = stream.Read(buf, 0, 1024);
                while (len > 0)
                {
                    fs.Write(buf, 0, len);
                    len = stream.Read(buf, 0, 1024);
                }
                stream.Close();
                fs.Close();
                return path;

            }
            catch (Exception ex)
            {

                throw;
            }
        }
    }
}
