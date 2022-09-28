using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using WDSE;
using WDSE.Decorators;
using WDSE.ScreenshotMaker;
using System.Drawing;
using System.Net.Http;
using System.IO;
using Newtonsoft.Json;
using RestSharp;
using System.Net;
using Microsoft.Graph;
using System.Text;
using IronXL;
using Google.Apis.Util.Store;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using System.Data.SqlClient;
using System.Configuration;
using Telegram.Bot;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Args;
using MySql.Data.MySqlClient;
using System.Data;
using Microsoft.Reporting.WebForms;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Net.Mime;
using Microsoft.Graph.Models;
using Telegram.Bot.Types;
using Google.Protobuf;

namespace WorkerEmail
{
    public class Worker : BackgroundService
    {
        //tutorial google drive
        //https://www.youtube.com/watch?v=pHOweM1Gl6c
        //create project
        //open api drive

        private readonly ILogger<Worker> _logger;
        static TelegramBotClient Bot = new TelegramBotClient("5278461864:AAGwGEV3aJx8ZFKFfRaFB6TMc4mOem-uoc8");
        public static String connectionString = "Data Source=KBIDRC-TIMAH-DBMS.ptkbi.com;Initial Catalog=TIN_KBI;Persist Security Info=True;User ID=dbapp;Password=P@ssw0rd2022";
        private const string PathToCredentials = @"E:\ServiceDariTimahDBMS\ServiceTINOperationalTelegram\credentials.json";

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }
        public override Task StartAsync(CancellationToken cancellationToken)
        {
            return base.StartAsync(cancellationToken);
        }
        public override Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Service stopped");
            return base.StopAsync(cancellationToken);
        }
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            try
            {
                Bot.StartReceiving(Array.Empty<UpdateType>());
                Bot.OnUpdate += BotOnUpdateReceived;
            }
            catch (Exception x)
            {
                SendMessage("6289630870658@c.us", "Tradeprogress fail " + x.Message);
            }
        }
        private static async void BotOnUpdateReceived(object sender, UpdateEventArgs e)
        {

            var message = e.Update.Message;

            if (message.Type == MessageType.Text)
            {
                var text = message.Text;

                if (message.ReplyToMessage != null)
                {
                    var msg = message.ReplyToMessage.Text;
                    //filter untuk approve shipping instructions
                    if (msg.Contains("*shipping Instructions*") && text.ToLower() == "approve")
                    {
                        monitoringServices("DOP_TinTelegram", "Approve Shipping Instructions", "10.10.10.99", "Live");
                        approveShippingInstruction(message.Chat.Id, msg);
                    }
                    //filter untuk insert withdrawal timah
                    if (msg.Contains("#withdrawal#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Insert withdrawal timah", "10.10.10.99", "Live");

                        insertWithdrawal(message.Chat.Id, msg, text);
                    }
                    //filter untuk check code member
                    if (msg.Contains("#check code#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Check code clearing member", "10.10.10.99", "Live");

                        checkmembercode(message.Chat.Id, text);
                    }
                    if (msg.Contains("#trade register#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Get trade register timah", "10.10.10.99", "Live");

                        reportTradeRegister(message.Chat.Id, text);
                    }
                    if (msg.Contains("#invoice timah dalam negeri#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Get invoice dalam negeri timah", "10.10.10.99", "Live");

                        invoiceTINDalamNegeri(message.Chat.Id, text);
                    }
                    if (msg.Contains("#input ceiling price#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Input ceiling price di telegram", "10.10.10.99", "Live");

                        inputCeilingPrice(message.Chat.Id, text);
                    }
                    if (msg.Contains("#generate dfs#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Generate DFS timah di telegram", "10.10.10.99", "Live");

                        generateDFS(message.Chat.Id, text);
                    }
                    if (msg.Contains("#generate invoice#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Generate invoice luar negeri timah di telegram", "10.10.10.99", "Live");

                        generateInvoice(message.Chat.Id, text);
                    }
                    if (msg.Contains("#invoice emas off#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Generate invoice emas off exchange timah di telegram", "10.10.10.99", "Live");

                        generateInvoiceemasoff(message.Chat.Id, text);
                    }
                    if (msg.Contains("#uspcalculatefeeresiprokal#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Generate report usp calculate fee resiprokal timah di telegram", "10.10.10.99", "Live");

                        uspcalculatefeeresiprokal(message.Chat.Id, text);
                    }
                    if (msg.Contains("#searchsuretybond#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Get data surety bond timah di telegram", "10.10.10.99", "Live");

                        searchSuretyBond(message.Chat.Id, text);
                    }
                    if (msg.Contains("#surentybond#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Update suretybond timah di telegram", "10.10.10.99", "Live");
                        updateSuretyBond(message.Chat.Id, text, msg);
                    }
                    if (msg.Contains("#requestdailystatementemas#"))
                    {
                        monitoringServices("DOP_TinTelegram", "Get report daily statement emas di telegram", "10.10.10.99", "Live");
                        getdailystatementemas(message.Chat.Id, text, msg);
                    }
                }
                if (text == "/cektradefeed@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "cek tradefeed timah di telegram", "10.10.10.99", "Live");
                    cekTradefeed(message.Chat.Id);
                }
                if (text == "/eod@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "EOD timah di telegram", "10.10.10.99", "Live");
                    eodProcess(message.Chat.Id);
                }
                if (text == "/sod@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "SOD timah di telegram", "10.10.10.99", "Live");
                    sodProcess(message.Chat.Id);
                }
                if (text == "/upserttks@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "EXEC SP upsert tks timah di telegram", "10.10.10.99", "Live");
                    upserttks(message.Chat.Id);
                }
                if (text == "/rolloverbgr@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "EXEC SP rollover bgr timah di telegram", "10.10.10.99", "Live");
                    rolloverbgr(message.Chat.Id);
                }
                if (text == "/invoicedalamnegeri@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Request invoice dalam negeri di telegram", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#invoice timah dalam negeri# Please reply to the message\n{yyyy-mm-dd}" + "_", ParseMode.Markdown);
                }
                if (text == "/financialinfokbi@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Get financial info KBI di telegram", "10.10.10.99", "Live");
                    financialInfoKbi(message.Chat.Id);
                }
                if (text == "/financialinfobbj@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Get financial info BBJ di telegram", "10.10.10.99", "Live");
                    financialInfoBbj(message.Chat.Id);
                }
                if (text == "/cekcodemember@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Request code member di telegram", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#check code# Please reply to the message with the member name" + "_", ParseMode.Markdown);
                }
                if (text == "/rpttraderegister@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Request trade register di telegram", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#trade register# Please reply to the message\n{yyyy-mm-dd}\n{seller code}\n{buyer code}" + "_", ParseMode.Markdown);
                }
                if (text == "/publishreport@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Publish report timah di telegram", "10.10.10.99", "Live");
                    publishreport(message.Chat.Id);
                }
                if (text == "/insertceilingprice@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Request insert ceiling price timah di telegram", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#input ceiling price# Please reply to the message with number" + "_", ParseMode.Markdown);
                }
                if (text == "/generatedfs@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Request generate dfs di telegram", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#generate dfs# Please reply to the message \n{yyyy-mm-dd}" + "_", ParseMode.Markdown);
                }
                if (text == "/generateinvoice@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Request generate invoice luar negeri di telegram", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#generate invoice# Please reply to the message \n{yyyy-mm-dd}" + "_", ParseMode.Markdown);
                }
                if (text == "/invoiceemasoff@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Request invoice emas off exchange di telegram", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#invoice emas off# Please reply to the message\n{yyyy-mm-dd} datestart\n{yyyy-mm-dd} dateend\n{vendorcode}\n====\nplm = pluangmas\nse = sakuemas" + "_", ParseMode.Markdown);
                }
                if (text == "/calculatefeeresiprokal@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Request usp calculate fee resiprokal di telegram", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#uspcalculatefeeresiprokal# Please reply to the message\n{yyyy-mm-dd} datestart\n{yyyy-mm-dd} dateend\n{yyyy-mm-dd} invoice date" + "_", ParseMode.Markdown);
                }
                if (text == "/rekapreporttimah@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Get rekap report timah di telegram", "10.10.10.99", "Live");
                    rekapReportTimah(message.Chat.Id);
                }
                if (text == "/updatesuretybond@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Request search suretybond timah di telegram", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#searchsuretybond# Please reply to the message\n{member code}" + "_", ParseMode.Markdown);
                }
                if (text == "/sendnotapemberitahuan@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Send email nota pemberitahuan", "10.10.10.99", "Live");
                    sendemailnotapemberitahuan(message.Chat.Id);
                }
                if (text == "/dailystatementemas@Timah_Bot")
                {
                    monitoringServices("DOP_TinTelegram", "Request daily statement emas off di telegram", "10.10.10.99", "Live");
                    Bot.SendTextMessageAsync(message.Chat.Id, "_#requestdailystatementemas# Please reply to the message\nstartdate {yyyy-MM-dd}\nenddate {yyyy-MM-dd}\nvendorcode {150/152}" + "_", ParseMode.Markdown);
                }
            }
            if (message.Type == MessageType.Document)
            {
                //download all document
                var file = await Bot.GetFileAsync(message.Document.FileId);
                FileStream fs = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\" + message.Document.FileName, FileMode.Create);
                await Bot.DownloadFileAsync(file.FilePath, fs);
                fs.Close();
                fs.Dispose();
                //filer disini untuk insert bukti tf
                if (message.Document.FileName.ToLower().Contains("acc_statement"))
                {
                    monitoringServices("DOP_TinTelegram", "Upload bukti transfer withdrawal timah di telegram", "10.10.10.99", "Live");
                    insertBuktiTFWithdrawal(message.Chat.Id, fs.Name);
                }
            }
        }
        public static void cekTradefeed(long chat_id)
        {
            try
            {
                var dr_1 = new DataSet1TableAdapters.RawTradefeedTableAdapter();
                var dr_2 = new DataSet2TableAdapters.DataTable1TableAdapter();
                var dt_kbi = dr_1.GetData(DateTime.Now.Date);
                var dt_bbj = dr_2.GetData(DateTime.Now.Date);
                Bot.SendTextMessageAsync(chat_id, "_Date: " + DateTime.Now.ToString("dd MMM yyyy") + "\n\nTotal Trade KBI: " + dt_kbi[0].Qty + "\nTotal Trade BBJ: " + dt_kbi[0].Qty + "\n\nTimeStamp: " + DateTime.Now.ToString("hh:mm:ss") + "_", ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Eror cek tradefeed " + ex.Message + "_", ParseMode.Markdown);
            }
        }
        public static void eodProcess(long chat_id)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_EOD Tanggal : " + DateTime.Now.ToString("dd-MMM-yyyy") + "\nProses EOD TIN start " + DateTime.Now.ToString("hh:mm:ss") + "_", ParseMode.Markdown);
                string queryStringEOD = "EXEC [SKD].[EOD_ProcessEOD] '" + DateTime.Now.Date.ToString("yyyy-MM-dd") + "','Robot KBI','10.12.1.60'";

                // WriteToFile(queryStringEOD);
                using (SqlConnection connection = new SqlConnection(
                           connectionString))
                {
                    SqlCommand command = new SqlCommand(queryStringEOD, connection);
                    connection.Open();
                    command.CommandTimeout = 1800;
                    SqlDataReader reader = command.ExecuteReader();
                    connection.Dispose();
                    connection.Close();
                }
                Bot.SendTextMessageAsync(chat_id, "_EOD Tanggal : " + DateTime.Now.ToString("dd-MMM-yyyy") + "\nProses EOD TIN selesai " + DateTime.Now.ToString("hh:mm:ss") + "_", ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Eror cek tradefeed " + ex.Message + "_", ParseMode.Markdown);

            }

        }
        public static void sodProcess(long chat_id)
        {
            var dr_sod = new DataSet1TableAdapters.QueriesTableAdapter();
            DateTime dateValue = DateTime.Now.Date;
            Bot.SendTextMessageAsync(chat_id, "_SOD Tanggal: " + DateTime.Now.ToString("dd - MMM - yyyy") + "\nProses SOD TIN start " + DateTime.Now.ToString("hh: mm:ss") + "_", ParseMode.Markdown);

            dr_sod.uspProcessSOD(dateValue, "Robot KBI", "10.12.1.60");
            var dr_param = new DataSet1TableAdapters.ParameterTableAdapter();
            var dt_BusinessDate = dr_param.GetDataByCode("BusinessDate");
            dt_BusinessDate[0].DateValue = dateValue;
            dr_param.Update(dt_BusinessDate);

            var dt_LastEOD = dr_param.GetDataByCode("LastEOD");
            dt_BusinessDate[0].DateValue = dateValue;
            dr_param.Update(dt_LastEOD);

            Bot.SendTextMessageAsync(chat_id, "_SOD Date Value : " + dateValue.ToString("dd MMM yyyy") + "_", ParseMode.Markdown);
            Bot.SendTextMessageAsync(chat_id, "_SOD Tanggal: " + DateTime.Now.ToString("dd - MMM - yyyy") + "\nProses SOD TIN selesai " + DateTime.Now.ToString("hh: mm:ss") + "_", ParseMode.Markdown);
        }
        public static void upserttks(long chat_id)
        {
            Bot.SendTextMessageAsync(chat_id, "_Prosess Upsert Tks Start " + DateTime.Now.ToString("hh:mm:ss") + "_", ParseMode.Markdown);


            string queryString = "EXEC [dbo].[UpsertStagingSellerAllocationTKS]";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlCommand command = new SqlCommand(queryString, connection);
                    connection.Open();
                    command.CommandTimeout = 1800;
                    SqlDataReader reader = command.ExecuteReader();
                    connection.Dispose();
                    connection.Close();
                }

                sodProcess(chat_id);
            }
            catch (Exception x)
            {
                Bot.SendTextMessageAsync(chat_id, "_Prosess Upsert Tks gagal " + x.Message + " " + DateTime.Now.ToString("hh:mm:ss") + "_", ParseMode.Markdown);
            }
            Bot.SendTextMessageAsync(chat_id, "_Prosess Upsert Tks Sukses " + DateTime.Now.ToString("hh:mm:ss") + "_", ParseMode.Markdown);
        }
        public static void rolloverbgr(long chat_id)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_Prosess RollOver BGR Start " + DateTime.Now.ToString("hh:mm:ss") + "_", ParseMode.Markdown);

                MySql.Data.MySqlClient.MySqlConnection dbConn = new MySql.Data.MySqlClient.MySqlConnection("Persist Security Info=False;server=10.15.2.60;database=bgrkbi;uid=hasto;password=Jakarta01");

                MySqlCommand cmd1 = new MySqlCommand("prc_rollover_ctd", dbConn);
                cmd1.CommandType = CommandType.StoredProcedure;
                cmd1.Connection.Open();
                cmd1.ExecuteNonQuery();
                cmd1.Connection.Close();
                Bot.SendTextMessageAsync(chat_id, "_Prosess RollOver BGR Sukses " + DateTime.Now.ToString("hh:mm:ss") + "_", ParseMode.Markdown);
                sodProcess(chat_id);
            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Prosess RollOver BGR gagal " + ex.Message + " " + DateTime.Now.ToString("hh:mm:ss") + "_", ParseMode.Markdown);
            }
        }
        public static async void invoiceTINDalamNegeri(long chat_id, string msg)
        {
            try
            {
                var data = msg.Split(' ');
                Bot.SendTextMessageAsync(chat_id, "_Generate start " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);

                string queryStringEOD = "EXEC [SKD].[EOD_GenerateTradeProgress] '" + msg + "','Robot KBI'";
                Bot.SendTextMessageAsync(chat_id, queryStringEOD);
                // WriteToFile(queryStringEOD);
                using (SqlConnection connection = new SqlConnection(
                           connectionString))
                {
                    SqlCommand command = new SqlCommand(queryStringEOD, connection);
                    connection.Open();
                    command.CommandTimeout = 1800;
                    SqlDataReader reader = command.ExecuteReader();
                    connection.Dispose();
                    connection.Close();
                }
                var context = new Data();
                Bot.SendTextMessageAsync(chat_id, "2");
                var data_trade = (from a in context.TradeProgresses where a.businessdate == Convert.ToDateTime(msg) select a).ToList();
                Bot.SendTextMessageAsync(chat_id, data_trade.Count.ToString());
                var clearingmember = (from a in context.clearingMembers where a.StatusDomisiliFlag == "D" select a).ToList();
                Bot.SendTextMessageAsync(chat_id, clearingmember.Count.ToString());
                //looping for trade progress
                foreach (var item in data_trade)
                {
                    //looping for check sellerId dalam negeri
                    foreach (var dt in clearingmember)
                    {
                        if (item.sellerId.Substring(0, 4) == dt.code)
                        {
                            //looping for buyer dalam negeri
                            foreach (var dt2 in clearingmember)
                            {
                                if (item.buyerId.Substring(0, 4) == dt2.code)
                                {
                                    Console.WriteLine("Masuk");
                                    string path = getReportSSRSWord("Invoice_DN", "&PeriodStart=" + data[0] + "&SellerId=" + item.sellerId + "&BuyerId=" + item.buyerId, "Invoice Dalam Negeri " + item.sellerId, "TIN_EOD_Report");
                                    sendFileTelegram(chat_id.ToString(), path);
                                    //string urlContents = await uploadGoogle(path, "Invoice TIN Dalam Negeri.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                                    //Bot.SendTextMessageAsync(chat_id, "https://drive.google.com/file/d/" + urlContents + "/view?usp=drivesdk \n\n" + DateTime.Now.ToString("HH:mm:ss"), ParseMode.Markdown);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "Generate incoive fail : " + ex.Message + " \n\n" + DateTime.Now.ToString("HH:mm:ss"), ParseMode.Markdown);
            }
        }
        public static void approveShippingInstruction(long chat_id, string message)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_Processing approve SI start\n" + DateTime.Now.ToString("hh: mm:ss") + "_", ParseMode.Markdown);

                string[] quotes = message.Split('*');
                var dr = new DataSet1TableAdapters.TradeFeedTableAdapter();
                var dt = dr.GetDataByNoSi(quotes[3]);

                for (int k = 0; k < dt.Count; k++)
                {
                    dt[k].ShippingInstructionFlag = "A";
                    dt[k].ShippingInstructionApproveDate = DateTime.Now;
                    dt[k].ApprovalDesc = "ok";
                    dt[k].LastUpdatedBy = "Robot KBI";
                    dt[k].LastUpdatedDate = DateTime.Now;
                    dr.Update(dt);
                }

                sendNOS(chat_id, quotes[3]);

            }
            catch (Exception xx)
            {
                Bot.SendTextMessageAsync(chat_id, "_Fail approve SI Time: " + xx.Message + "\n" + DateTime.Now.ToString("hh: mm:ss") + "_", ParseMode.Markdown);
            }
        }
        public static void sendNOS(long chat_id, string noSi)
        {
            try
            {
                string fileName = noSi;
                if (noSi.Contains("/"))
                {
                    fileName = noSi.Replace("/", "_");
                }
                if (fileName.Contains(":"))
                {
                    fileName = fileName.Replace(":", "_");
                }

                string path = getReportSSRSPDF("RptNoticeOfShipment", "&NoSI=" + noSi, fileName);

                sendFileTelegram("-1001649045625", path);

                sendEmailNOS(path, fileName);
                Bot.SendTextMessageAsync(chat_id, "Success send email Notice Of Shipment\n" + DateTime.Now.ToString("hh:mm:ss"), ParseMode.Markdown);
                System.IO.File.Delete(path);
            }
            catch (Exception x)
            {
                Bot.SendTextMessageAsync(chat_id, "_Send NOS failed : " + x.Message + "\n" + DateTime.Now.ToString("hh: mm:ss") + "_", ParseMode.Markdown);
            }
        }
        public static void sendEmailNOS(string filePath, string filename)
        {
            var stream = System.IO.File.Open(filePath, FileMode.Open);
            //var x = new Attachment(stream);


            //send email
            Spire.Email.MailAddress addressFrom = "automatic_ptkbi@outlook.com";

            Spire.Email.MailAddress[] addressTo = new Spire.Email.MailAddress[4];
            addressTo[0] = "drp@ptkbi.com";
            addressTo[1] = "wiridon92@gmail.com";
            addressTo[2] = "yogi.kusumas@bgrindonesia.co.id";
            addressTo[3] = "tantrakaryasejahtera@gmail.com";


            Spire.Email.MailMessage message1 = new Spire.Email.MailMessage(addressFrom, addressTo);
            //message1.Cc = "jujuk1020@gmail.com";
            //message1.Cc = "drp@ptkbi.com";
            message1.Subject = "Notice Of Shipment";
            message1.BodyText = "\nYth Bapak Ibu\nPT BGR dan PT TKS,\r\n\n" + "Berikut kami kirimkan Notice Of Shipment dengan nomor Shipping Instructions : " + filename + "\r\n\n***Pesan ini di buat secara otomatis***\r\n\n" + "Terima kasih\nPT.KBI";
            message1.Attachments.Add(new Spire.Email.Attachment(stream, filename + ".pdf"));

            Spire.Email.Smtp.SmtpClient smtp = new Spire.Email.Smtp.SmtpClient();
            smtp.Host = "smtp.outlook.com";
            smtp.ConnectionProtocols = Spire.Email.IMap.ConnectionProtocols.Ssl;
            smtp.Username = addressFrom.Address;
            smtp.Password = "Jakarta2021";
            smtp.Port = 587;
            smtp.SendOne(message1);
            stream.Close();
            //x.Dispose();
        }
        public static async void financialInfoKbi(long chat_id)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_Generate start" + "_", ParseMode.Markdown);
                string path = getReportSSRSPDF("RptFinancialInfoForWhatsapp", "&businessDate=" + DateTime.Now.ToString("yyyy-MM-dd"), "Financial Info KBI");
                sendFileTelegram(chat_id.ToString(), path);
                Bot.SendTextMessageAsync(chat_id, "_Generate finish" + "_", ParseMode.Markdown);

                //string urlContents = await uploadGoogle(path, "financial info kbi.pdf", "application/x-pdf");
                //Bot.SendTextMessageAsync(chat_id, "_https://drive.google.com/file/d/" + urlContents + "/view?usp=drivesdk" + "_", ParseMode.Markdown);
            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_hGenerate fail " + ex.Message + "_", ParseMode.Markdown);
            }

        }
        public static async void financialInfoBbj(long chat_id)
        {
            try
            {
                string path = getReportSSRSPDF("RptFinancialInfoStaging", "&businessDate=" + DateTime.Now.ToString("yyyy-MM-dd"), "Financial Info BBJ");
                Bot.SendTextMessageAsync(chat_id, "_Generate start" + "_", ParseMode.Markdown);
                sendFileTelegram(chat_id.ToString(), path);
                Bot.SendTextMessageAsync(chat_id, "_Generate finish" + "_", ParseMode.Markdown);

                //string urlContents = await uploadGoogle(path, "financial info bbj.pdf", "application/x-pdf");
                //Bot.SendTextMessageAsync(chat_id, "_https://drive.google.com/file/d/" + urlContents + "/view?usp=drivesdk" + "_", ParseMode.Markdown);
            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Generate fail " + ex.Message + "_", ParseMode.Markdown);
            }
        }
        public static void insertBuktiTFWithdrawal(long chat_id, string path)
        {
            try
            {
                bool format_old = false;
                List<data_bukti_transfer> dataCsv = new List<data_bukti_transfer>();
                string[] arrBuktiTf = System.IO.File.ReadAllLines(path);
                int j = 0;
                if (arrBuktiTf[0].Contains(";"))
                {
                    for (int i = 1; i < arrBuktiTf.Length; i++)
                    {
                        string[] arrDataTf = arrBuktiTf[i].Split(';');
                        dataCsv.Add(new data_bukti_transfer
                        {
                            AccountNo = (arrDataTf[0].ToString()),
                            Date = (arrDataTf[2].ToString().Remove(arrDataTf[2].Length - 9)),
                            ValDate = (arrDataTf[2].ToString().Remove(arrDataTf[2].Length - 9)),
                            Description1 = (arrDataTf[3].ToString()),
                            Description2 = (arrDataTf[4].ToString()),
                            Debit = (arrDataTf[6].ToString()),
                            Credit = (arrDataTf[5].ToString()),
                        });
                    }
                }
                else
                {
                    format_old = true;
                    StreamReader sr = new StreamReader(path);
                    while (!sr.EndOfStream)
                    {
                        MatchCollection matches = new Regex("((?<=\")[^\"]*(?=\"(,|$)+)|(?<=,|^)[^,\"]*(?=,|$))").Matches(sr.ReadLine());
                        string[] rows = new string[matches.Count + 1];
                        int i = 1;
                        foreach (var item in matches)
                        {
                            rows[i] = item.ToString();
                            i++;
                        }

                        if (j >= 1)
                        {
                            dataCsv.Add(new data_bukti_transfer
                            {
                                AccountNo = (rows[1].ToString()),
                                Date = (rows[2].ToString()),
                                ValDate = (rows[3].ToString()),
                                TransactionCode = (rows[4].ToString()),
                                Description1 = (rows[5].ToString()),
                                Description2 = (rows[6].ToString()),
                                ReferenceNo = (rows[7].ToString()),
                                Debit = (rows[8].ToString()),
                                Credit = (rows[9].ToString()),
                            });
                        }
                        j++;
                    }
                }

                var dt = new DataSet1TableAdapters.Bank_TransferTableAdapter();

                foreach (var item in dataCsv)
                {
                    if (item.Credit != ".00")
                    {
                        Decimal number_credit = Convert.ToDecimal(item.Credit.Replace(",", string.Empty).Replace(".", ","));
                        DateTime date = new DateTime();
                        var date1 = DateTime.Now.ToString("dd MMMM yyyy");
                        if (format_old == true)
                        {
                            date = DateTime.ParseExact(item.Date, "dd/MM/yy", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            date = DateTime.ParseExact(item.Date, "dd MMMM yyyy", CultureInfo.InvariantCulture);
                        }
                        var dr = dt.GetDataByFilter(date, item.Description2, number_credit);

                        if (dr.Count == 0)
                        {
                            int newId = -1;
                            Object retObj = null;
                            retObj = dt.InsertQuery(date, item.Description1, item.Description2, number_credit, DateTime.Now, "Robot KBI");
                            newId = Convert.ToInt32(retObj);

                            Bot.SendTextMessageAsync(chat_id, "_" + newId + "#withdrawal#\nThis Amount Transfer : " + number_credit + "\nPlease answer this message with the format :\n1.ExchangeReff\n2.Amount\n3.Use Secfund(Yes/No)\n\nTimeStamp: " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);

                            //SendMessage(chat_id, newId + "#withdrawal#\nThis Amount Transfer : " + number_credit + "\nPlease answer this message with the format :\n1.ExchangeReff\n2.Amount\n3.Use Secfund(Yes/No)\n\nTimeStamp: " + DateTime.Now.ToString("HH:mm:ss"));
                            //SendTelegram("-1001641017183", "[" + DateTime.Now.ToString("dd MMM yyyy") + "] [" + DateTime.Now.ToString("hh:mm:ss") + "]\nLog After Parsing: Success Insert ID Transfer = " + newId);

                        }
                        else
                        {
                            Bot.SendTextMessageAsync(chat_id, "This data has been uploaded " + DateTime.Now.ToString("HH:mm:ss"));
                        }
                    }
                }
                Bot.SendTextMessageAsync(chat_id, "_Processing upload proof of payment success \n " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Processing upload proof of payment fail " + ex.Message + "\n" + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
            }
        }
        public static void insertWithdrawal(long chat_id, string msg, string msgBody)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_Processing insert withdrawal start " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                string[] body = msgBody.Split('\n');
                string[] quotes = msg.Split('#');
                var dr = new DataSet1TableAdapters.uspTradeProgressByExchangeReffTableAdapter();
                var drw = new DataSet1TableAdapters.Rpt_WithdrawalTableAdapter();
                var drBT = new DataSet1TableAdapters.Bank_TransferTableAdapter();
                var drCM = new DataSet1TableAdapters.DataTable1TableAdapter();
                var dr_CP = new DataSet1TableAdapters.CMProfileTableAdapter();
                if (body[0].Contains("-"))
                {
                    List<dataInsertWD> data = new List<dataInsertWD>();
                    var nominal = body[1].Split('-');
                    var excreff = body[0].Split('-');
                    var secfundData = body[2].Split('-');
                    string sellercode = "";
                    string sellername = "";
                    for (int i = 0; i < nominal.Length; i++)
                    {
                        data.Add(new dataInsertWD
                        {
                            ammount = nominal[i],
                            exchangeReff = excreff[i],
                            secfund = secfundData[i]
                        });
                    }
                    foreach (var item in data)
                    {
                        var dt = dr.GetData(item.exchangeReff);
                        var dtBT = drBT.GetDataById(Convert.ToInt32(quotes[0]));
                        var dtCM = drCM.GetData(dt[0].sellerId);
                        var dt_CP = dr_CP.GetDataByCode(dtCM[0].CMCode);
                        Decimal secfund = 0;
                        if (item.secfund.ToLower() == "yes")
                        {
                            secfund = dt[0].secfund;
                        }

                        drw.Insert(dt[0].BusinessDate, dtCM[0].CMCode, dtCM[0].seller, dt_CP[0].CMBankName, dt_CP[0].CMAccountNo, dt[0].ExchangeRef, dt[0].ammount, secfund, Convert.ToDecimal(item.ammount), Convert.ToInt32(quotes[0]), dt[0].BusinessDate, 0, dt[0].buyerId, dt[0].sellerId);
                        sellercode = dtCM[0].CMCode;
                        sellername = dtCM[0].seller;
                    }
                    string path = getReportSSRSPDF("RptWithdrawal", "&BusinessDate=*&exchangeReff=-&seller=" + sellercode + "&id_bank_transfer=" + quotes[0], "Withdrawal");
                    sendFileTelegram(chat_id.ToString(), path);
                }
                else
                {
                    var dt = dr.GetData(body[0]);

                    var dtBT = drBT.GetDataById(Convert.ToInt32(quotes[0]));

                    var dtCM = drCM.GetData(dt[0].sellerId);

                    var dt_CP = dr_CP.GetDataByCode(dtCM[0].CMCode);

                    Decimal secfund = 0;
                    if (body[2].ToLower() == "yes")
                    {
                        secfund = dt[0].secfund;
                    }
                    drw.Insert(dt[0].BusinessDate, dtCM[0].CMCode, dtCM[0].seller, dt_CP[0].CMBankName, dt_CP[0].CMAccountNo, dt[0].ExchangeRef, dt[0].ammount, secfund, Convert.ToDecimal(body[1]), Convert.ToInt32(quotes[0]), dt[0].BusinessDate, 0, dt[0].buyerId, dt[0].sellerId);
                    string path = getReportSSRSPDF("RptWithdrawal", "&BusinessDate=-&exchangeReff=" + dt[0].ExchangeRef + "&seller=" + dtCM[0].CMCode + "&id_bank_transfer=" + quotes[0], "Widhrawal");
                    sendFileTelegram(chat_id.ToString(), path);
                    Bot.SendTextMessageAsync(chat_id, "_Generate withdrawal finish" + "_", ParseMode.Markdown);

                }

            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Processing insert withdrawal Fail " + ex.Message + " " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);

            }
        }
        public static void checkmembercode(long chat_id, string msg)
        {
            try
            {
                var dr = new DataSet1TableAdapters.CMProfileTableAdapter();
                var dt = dr.GetCodeByName(msg);
                foreach (var item in dt)
                {
                    Bot.SendTextMessageAsync(chat_id, "_" + item.Name + " - " + item.Code + "_", ParseMode.Markdown);
                }
            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Get member code fail " + ex.Message + "_", ParseMode.Markdown);

            }
        }
        public static void reportTradeRegister(long chat_id, string msg)
        {
            try
            {
                string[] data = msg.Split('\n');
                string path = getReportSSRSPDF("RptEODTradeRegisterForWA", "&businessDate=" + data[0] + "&clearingMemberId=" + data[1] + "&codeSeller=" + data[2], "Trade Register");
                sendFileTelegram(chat_id.ToString(), path);
            }
            catch (Exception)
            {

                throw;
            }
        }
        public static void publishreport(long chat_id)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_Publish report start" + "_", ParseMode.Markdown);

                IWebDriver ChromeDriver = new ChromeDriver();
                ChromeDriver.Navigate().GoToUrl("http://septi.ptkbi.com/login.aspx?ReturnUrl=%2fDefault.aspx");

                IWebElement username = ChromeDriver.FindElement(By.Id("uiAuthLogin_UserName"));
                IWebElement password = ChromeDriver.FindElement(By.Id("uiAuthLogin_Password"));
                IWebElement submit_login = ChromeDriver.FindElement(By.Id("uiAuthLogin_LoginImageButton"));

                username.SendKeys("Hasto");
                password.SendKeys("Jakarta2019");
                submit_login.Click();
                ChromeDriver.Navigate().GoToUrl("http://septi.ptkbi.com/ClearingAndSettlement/PublishReport2.aspx");

                //IWebElement Date = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_CtlCalendarPickUpEOD_uiTxtCalendar"));
                IWebElement Date = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_CtlCalendarPickUpEOD_uiTxtCalendar"));
                IWebElement submit_publish = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_uiBtnPublish"));

                IJavaScriptExecutor js = (IJavaScriptExecutor)ChromeDriver;

                js.ExecuteScript("arguments[0].removeAttribute('readonly','readonly')", Date);

                Date.Clear();
                var tgl_eod = DateTime.Now.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture);
                //var tgl_eod = "01-Jan-2020";
                Date.SendKeys(tgl_eod);
                IWebElement loading = ChromeDriver.FindElement(By.Id("ctl00_ContentPlaceHolder1_UpdateProgress1"));
                submit_publish.Submit();
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                if (loading.Displayed)
                {
                    Thread.Sleep(60000);
                }
                ChromeDriver.Quit();
                Bot.SendTextMessageAsync(chat_id, "_Publish report finish" + "_", ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Publish report failed" + ex.Message + "_", ParseMode.Markdown);

            }
        }
        public static void inputCeilingPrice(long chat_id, string msg)
        {
            try
            {
                var drc = new DataSet1TableAdapters.CeilingPriceTableAdapter();
                var dtc = drc.GetData(13);
                dtc[0].FloorPrice = Convert.ToDecimal(msg);
                dtc[0].LastUpdatedDate = DateTime.Now;
                dtc[0].LastUpdatedBy = "Robot KBI";
                drc.Update(dtc);
                Bot.SendTextMessageAsync(chat_id, "_update ceiling price succses, value : " + msg + "_", ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_update ceiling price failed : " + ex.Message + "_", ParseMode.Markdown);

            }
        }
        public static void generateDFS(long chat_id, string msg)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_generate dfs start " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                string path = getReportSSRSPDFExcel("uspRptEODDFS_bydate", "&businessDate=" + msg, "DFS", "TIN_EOD_Report");
                sendFileTelegram(chat_id.ToString(), path);
            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_generate dfs failed " + ex.Message + " " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
            }

        }
        public static void generateInvoice(long chat_id, string msg)
        {
            try
            {
                List<String> filePathAll = new List<string>();
                Bot.SendTextMessageAsync(chat_id, "_generate invoice start " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                DateTime tanggal = Convert.ToDateTime(msg);
                var dr = new DataSet1TableAdapters.EODTradeProgress2TableAdapter();
                var drBuyer = new DataSet1TableAdapters.EODTradeProgress1TableAdapter();
                var dr_cm = new DataSet1TableAdapters.ClearingMemberTableAdapter();
                var dt = dr.GetData(tanggal);
                var dt_buyerId = drBuyer.GetData(tanggal);
                if (dt.Count != 0)
                {
                    //looping untuk seller
                    foreach (var item in dt)
                    {
                        var dt_seller = dr_cm.GetDataByCode(item.SellerId.Substring(0, 4));

                        String SellerCMID = dt_seller[0].CMID.ToString();
                        String sellerName = dt_seller[0].Name.ToString();

                        string path = getReportSSRSWord("Invoice Seller", "&BusinessDate=" + msg + "&SellerCMID=" + SellerCMID, "Invoice Seller " + sellerName + " " + Convert.ToDateTime(msg).ToString("dd MMM yyyy"), "TIN_EOD_Report");

                        filePathAll.Add(path);

                    }
                }

                //LOOPING FOR BUYER
                if (dt_buyerId.Count != 0)
                {
                    //looping untuk buyer
                    foreach (var item in dt_buyerId)
                    {
                        //FOR BUYER

                        var dt_buyer = dr_cm.GetDataByCode(item.BuyerId.Substring(0, 4));
                        String BuyerCMID = dt_buyer[0].CMID.ToString();
                        String BuyerName = dt_buyer[0].Name.ToString();

                        string path = getReportSSRSWord("Invoice Buyer", "&BusinessDate=" + msg + "&BuyerCMID=" + BuyerCMID, "Invoice Buyer " + BuyerName + " " + Convert.ToDateTime(msg).ToString("dd MMM yyyy"), "TIN_EOD_Report");

                        filePathAll.Add(path);
                    }
                }
                foreach (var item in filePathAll)
                {
                    sendFileTelegram(chat_id.ToString(), item);
                }
            }
            catch (Exception x)
            {
                Bot.SendTextMessageAsync(chat_id, "_generate invoice failed " + x.Message + " " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
            }
        }
        public static async void generateInvoiceemasoff(long chat_id, string msg)
        {
            try
            {
                var data = msg.Split('\n');
                Bot.SendTextMessageAsync(chat_id, "_generate invoice start " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                string path = getReportSSRSPDFExcel("Rpt_InvoiceEmas", "&DateStart=" + data[0] + "&DateEnd=" + data[1] + "&VendorCode=" + data[2], "Invoice Emas Off.xlsx", "EmassOff");
                sendFileTelegram(chat_id.ToString(), path);
            }
            catch (Exception x)
            {
                Bot.SendTextMessageAsync(chat_id, "_generate invoice emas off eror " + x.Message + " " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);

            }
        }
        public static async void uspcalculatefeeresiprokal(long chat_id, string msg)
        {
            try
            {
                List<string> listPath = new List<string>();
                var data = msg.Split('\n');
                Bot.SendTextMessageAsync(chat_id, "_" + "Generate Report Start " + DateTime.Now.ToString("HH: mm:ss") + "_", ParseMode.Markdown);

                //TKS DALAM NEGERI
                string path = getReportSSRSWord("InvoiceCalculateFeeReceiptLocal", " &PeriodStart=" + data[0] + "&PeriodEnd=" + data[1] + "&InvoiceDate=" + data[2] + "&status=D&warehouse=TKS", "Invoice Calculate Receive Dalam Negeri TKS", "TIN_EOD_Report");
                string urlContents = await uploadGoogle(path, "Invoice Calculate Receive Dalam Negeri TKS", "application/msword");
                listPath.Add("Inv dalam negeri TKS\nhttps://drive.google.com/file/d/" + urlContents + "/view?usp=drivesdk\n\n");
                //TKS LUAR NEGERI
                path = getReportSSRSWord("InvoiceCalculateFeeReceiptLocal", " &PeriodStart=" + data[0] + "&PeriodEnd=" + data[1] + "&InvoiceDate=" + data[2] + "&status=L&warehouse=TKS", "Invoice Calculate Receive Luar Negeri TKS", "TIN_EOD_Report");
                urlContents = await uploadGoogle(path, "Invoice Calculate Receive Luar Negeri TKS", "application/msword");
                listPath.Add("Inv luar negeri TKS\nhttps://drive.google.com/file/d/" + urlContents + "/view?usp=drivesdk\n\n");
                //BGR DALAM NEGERI
                path = getReportSSRSWord("InvoiceCalculateFeeReceiptLocal", " &PeriodStart=" + data[0] + "&PeriodEnd=" + data[1] + "&InvoiceDate=" + data[2] + "&status=D&warehouse=BGR", "Invoice Calculate Receive Dalam Negeri BGR", "TIN_EOD_Report");
                urlContents = await uploadGoogle(path, "Invoice Calculate Receive Dalam Negeri BGR", "application/msword");
                listPath.Add("Inv dalam negeri BGR\nhttps://drive.google.com/file/d/" + urlContents + "/view?usp=drivesdk\n\n");
                //BGR LUAR NEGERI
                path = getReportSSRSWord("InvoiceCalculateFeeReceiptLocal", " &PeriodStart=" + data[0] + "&PeriodEnd=" + data[1] + "&InvoiceDate=" + data[2] + "&status=L&warehouse=BGR", "Invoice Calculate Receive LUAR Negeri BGR", "TIN_EOD_Report");
                urlContents = await uploadGoogle(path, "Invoice Calculate Receive Luar Negeri BGR", "application/msword");
                listPath.Add("Inv luar negeri BGR\nhttps://drive.google.com/file/d/" + urlContents + "/view?usp=drivesdk\n\n");

                Bot.SendTextMessageAsync(chat_id, String.Join("\n", listPath) + "\nProcessing success " + DateTime.Now.ToString("HH:mm:ss"), ParseMode.Markdown);

            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_Error get uspcalculatefee" + ex.Message + "\n\nProcessing success " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
            }
        }
        public static void rekapReportTimah(long chat_id)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_generate rekap report start " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                string path = getReportSSRSPDFExcel("RptDailyTrxTimah", "", "Rekap report timah.xlsx", "TIN_EOD_Report");
                sendFileTelegram(chat_id.ToString(), path);
            }
            catch (Exception x)
            {
                Bot.SendTextMessageAsync(chat_id, "_generate rekap report eror " + x.Message + " " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);

            }

        }
        public static void searchSuretyBond(long chat_id, string msg)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_Searching surety bond start " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);

                List<String> filePathAll = new List<string>();
                var dr = new DataSet3TableAdapters.InvestorTableAdapter();
                var dr_sb = new DataSet3TableAdapters.SuretyBondTableAdapter();
                var dt = dr.GetDataByCode(msg);
                if (dt.Count != 0)
                {
                    foreach (var item in dt)
                    {
                        var dt_sb = dr_sb.GetDataByInvId(item.InvestorID);
                        if (dt_sb.Count != 0)
                        {
                            foreach (var data in dt_sb)
                            {
                                Bot.SendTextMessageAsync(chat_id, "_" + data.SuretyBondID + "#surentybond#\nName : " + item.Name + "\n1.Bond Serial No : " + data.BondSerialNo + "\n2.Bond Ammount : " + data.Amount + "\n3.Remain Ammount : " + data.RemainAmount + "\n4.Exp Date : " + data.ExpireDate.ToString("yyyy-MM-dd") + "\n5.Exp Date Haircut : " + data.ExpDateHaircut.ToString("yyyy-mm-dd") + "\nPlease answer this message to update data\n{no}{.}{value} \nTime Send : " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                            }
                        }
                        else
                        {
                            //Bot.SendTextMessageAsync(chat_id, "Surety Bond is empty ! please check again");
                        }
                    }
                }
                else
                {
                    Bot.SendTextMessageAsync(chat_id, "Investor is empty ! please check again");
                }
                Bot.SendTextMessageAsync(chat_id, "_Searching surety bond finish " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);

            }
            catch (Exception x)
            {
                Bot.SendTextMessageAsync(chat_id, "Search surety bond fail " + x.Message);
            }
        }
        public static void updateSuretyBond(long chat_id, string msg, string quote)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_Update surety bond start " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                string[] quotes = quote.Split('#');
                var dr_sb = new DataSet3TableAdapters.SuretyBondTableAdapter();
                var spltmsg = msg.Split('\n');
                var dt_sb = dr_sb.GetDataByIdsb(Convert.ToInt32(quotes[0]));

                for (int i = 0; i < spltmsg.Length; i++)
                {
                    var type = spltmsg[i].Split('.');
                    if (type[0] == "1")
                    {
                        dt_sb[0].BondSerialNo = type[1];
                        dt_sb[0].LastUpdatedBy = "Robot KBI";
                        dt_sb[0].LastUpdatedDate = DateTime.Now;
                        dr_sb.Update(dt_sb);
                    }
                    if (type[0] == "2")
                    {
                        dt_sb[0].Amount = Convert.ToDecimal(type[1]);
                        dt_sb[0].LastUpdatedBy = "Robot KBI";
                        dt_sb[0].LastUpdatedDate = DateTime.Now;
                        dr_sb.Update(dt_sb);
                    }
                    if (type[0] == "3")
                    {
                        dt_sb[0].RemainAmount = Convert.ToDecimal(type[1]);
                        dt_sb[0].LastUpdatedBy = "Robot KBI";
                        dt_sb[0].LastUpdatedDate = DateTime.Now;
                        dr_sb.Update(dt_sb);
                    }
                    if (type[0] == "4")
                    {
                        dt_sb[0].ExpireDate = Convert.ToDateTime(type[1]);
                        dt_sb[0].LastUpdatedBy = "Robot KBI";
                        dt_sb[0].LastUpdatedDate = DateTime.Now;
                        dr_sb.Update(dt_sb);
                    }
                    if (type[0] == "5")
                    {
                        dt_sb[0].ExpDateHaircut = Convert.ToDateTime(type[1]);
                        dt_sb[0].LastUpdatedBy = "Robot KBI";
                        dt_sb[0].LastUpdatedDate = DateTime.Now;
                        dr_sb.Update(dt_sb);
                    }
                }
                Bot.SendTextMessageAsync(chat_id, "_Update surety bond finish " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
            }
            catch (Exception x)
            {
                Bot.SendTextMessageAsync(chat_id, "_Update surety bond fail " + x.Message + " " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
            }
        }
        public static void sendemailnotapemberitahuan(long chat_id)
        {
            try
            {
                Bot.SendTextMessageAsync(chat_id, "_send email start " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                var dt = new DataSet3TableAdapters.EODTradeProgressTableAdapter();
                var dr = dt.GetData();
                if (dr.Count != 0)
                {
                    foreach (var item in dr)
                    {
                        if (item.SellerId.Substring(0, 4) == "SMS3" || item.BuyerId.Substring(0, 4) == "BAL0")
                        {
                            string path = getReportSSRSPDF("RptEDONotaPemberitahuanRevision", "&businessDate="+item.BusinessDate.ToString("yyyy-MM-dd")+"&sellerid=" + item.SellerId.Substring(0, 4) + "&buyerid=" + item.BuyerId.Substring(0, 4), "Nota Pemberitahuan "+DateTime.Now.ToString("HH_mm_ss_sss"));
                            var stream1 = System.IO.File.Open(path, FileMode.Open);
                            string path2 = getReportSSRSPDF("RptEODTradeRegisterForWA", "&businessDate=" + item.BusinessDate.ToString("yyyy-MM-dd") + "&clearingMemberId=" + item.BuyerId.Substring(0, 4) + "&codeSeller=" + item.SellerId.Substring(0, 4), "Trade Register " + DateTime.Now.ToString("HH_mm_ss_sss"));
                            var stream2 = System.IO.File.Open(path2, FileMode.Open);
                            var dt_name = new DataSet3TableAdapters.ClearingMemberTableAdapter();
                            var dr_name = dt_name.GetDataByCode(item.BuyerId.Substring(0, 4));
                            var dr_name_seller = dt_name.GetDataByCode(item.SellerId.Substring(0, 4));
                            var dt_email = new DataSet3TableAdapters.EventRecipientListTableAdapter();
                            var dr_email_buyer = dt_email.GetDataByCode(item.BuyerId.Substring(0, 4));
                            var dr_email_seller = dt_email.GetDataByCode(item.SellerId.Substring(0, 4));
                            var dr_total = new DataSet3TableAdapters.DataTable1TableAdapter();
                            var dt_total = dr_total.GetData(item.BusinessDate, item.BuyerId.Substring(0, 4), item.SellerId.Substring(0, 4));

                            try
                            {
                                string file = AppDomain.CurrentDomain.BaseDirectory + "\\index.html";
                                string text = System.IO.File.ReadAllText(file);
                                text = text.Replace("#amount#", dt_total[0].Amount.ToString("#,##0"));
                                text = text.Replace("#businessdate#", item.BusinessDate.ToString("dd-MMM-yyyy"));
                                text = text.Replace("#name#", dr_name[0].Name);
                                text = text.Replace("#seller#", dr_name_seller[0].Name);
                                text = text.Replace("#nova#", dt_total[0].AccountNo);

                                string seting_email_file = AppDomain.CurrentDomain.BaseDirectory + "\\seting_email.txt";
                                string seting_email = System.IO.File.ReadAllText(seting_email_file);
                                string[] email_cc = seting_email.Split("\r\n");

                                ////send email
                                //Spire.Email.MailAddress addressFrom = "automatic_ptkbi@outlook.com";
                                //Spire.Email.MailAddress[] addressTo = new Spire.Email.MailAddress[dr_email_buyer.Count()];
                                //int index = 0;
                                //foreach (var item_send in dr_email_buyer)
                                //{
                                //    addressTo[index] = item_send.EmailAddress;
                                //    index++;
                                //}

                                //Spire.Email.MailMessage message1 = new Spire.Email.MailMessage(addressFrom, addressTo);
                                //message1.Cc.Add("drp@ptkbi.com");
                                //message1.Cc.Add("yunita.esaa@gmail.com");
                                //message1.Cc.Add("jujuk1020@gmail.com");
                                //foreach (var item_send in dr_email_seller)
                                //{
                                //    message1.Cc.Add(item_send.EmailAddress);
                                //}
                                //message1.Subject = "Payment Obligation for " + item.BusinessDate.ToString("dd-MM-yyyy") + " trading day";
                                //message1.BodyHtml = text;

                                //message1.Attachments.Add(new Spire.Email.Attachment(stream1, "Nota Pemberitahuan.pdf"));
                                //message1.Attachments.Add(new Spire.Email.Attachment(stream2, "Trade Register.pdf"));

                                //Spire.Email.Smtp.SmtpClient smtp = new Spire.Email.Smtp.SmtpClient();
                                //smtp.Host = "smtp.outlook.com";
                                //smtp.ConnectionProtocols = Spire.Email.IMap.ConnectionProtocols.Ssl;
                                //smtp.Username = addressFrom.Address;
                                //smtp.Password = "Jakarta2021";
                                //smtp.Port = 587;
                                //smtp.SendOne(message1);
                                //stream1.Close();
                                //stream2.Close();
                                ////end test


                                System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
                                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient();
                                message.From = new System.Net.Mail.MailAddress("pb@ptkbi.com");
                                foreach (var item_send in dr_email_buyer)
                                {
                                    message.To.Add(new System.Net.Mail.MailAddress(item_send.EmailAddress));
                                }

                                foreach (var item_send in dr_email_seller)
                                {
                                    message.CC.Add(new System.Net.Mail.MailAddress(item_send.EmailAddress));
                                }
                                foreach (var item_email in email_cc)
                                {
                                    message.CC.Add(new System.Net.Mail.MailAddress(item_email));
                                }

                                message.Subject = "Payment Obligation for " + item.BusinessDate.ToString("dd-MM-yyyy") + " trading day";
                                message.IsBodyHtml = true; //to make message body as html  
                                message.Body = text;
                                message.Attachments.Add(new System.Net.Mail.Attachment(stream1, "Nota Pemberitahuan.pdf"));
                                message.Attachments.Add(new System.Net.Mail.Attachment(stream2, "Trade Register.pdf"));
                                smtp.Port = 25;
                                smtp.Host = "10.10.10.2"; //for gmail host  
                                smtp.EnableSsl = true;
                                smtp.UseDefaultCredentials = false;
                                smtp.EnableSsl = false;
                                smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                                new Task(delegate
                                {
                                    smtp.Send(message);
                                }).Start();
                            }
                            catch (Exception ex)
                            {
                                monitoringServices("DOP_TinTelegram", "Send email nota pemberitahuan error "+ex.Message, "10.10.10.99", "Eror");

                            }
                            Bot.SendTextMessageAsync(chat_id, "_Success proccesing send email MSP atau Amalgamet " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);

                        }
                        else
                        {
                            Bot.SendTextMessageAsync(chat_id, "_Tidak ada transaksi Amalgamet atau MSP " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Bot.SendTextMessageAsync(chat_id, "_processing fail " +ex.Message+" "+ DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);

            }
        }
        public static void getdailystatementemas(long chat_id, string msg, string quote)
        {
            try
            {
                var data = msg.Split('\n');
                Bot.SendTextMessageAsync(chat_id, "_generate report start " + DateTime.Now.ToString("HH:mm:ss") + "_", ParseMode.Markdown);
                string path = getReportSSRSPDFExcel("RptDailyStatement", "&datestart=" + data[0] + "&dateend=" + data[1] + "&vendorcode=" + data[2], "Daily Statement Emas Off " + data[2], "EmassOff");
                sendFileTelegram(chat_id.ToString(), path);

            }
            catch (Exception ex)
            {
                monitoringServices("DOP_TinTelegram", "Get report daily statement emas di telegram eror : "+ex.Message, "10.10.10.99", "Eror");

                throw;
            }
        }
        public class dataInsertWD
        {
            public string exchangeReff { get; set; }
            public string ammount { get; set; }
            public string secfund { get; set; }
        }
        public static string getReportSSRSPDF(string filename, string parameter, string name)
        {
            try
            {
                string url = "http://10.15.2.60/ReportServerEOD?/TIN_EOD_Report/" + filename + "&rs:Command=Render&rs:Format=PDF&rc:OutputFormat=PDF" + parameter;


                System.Net.HttpWebRequest Req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                Req.Credentials = new NetworkCredential("Administrator", "Jakarta01");
                Req.Method = "GET";

                string path = AppDomain.CurrentDomain.BaseDirectory + "report\\" + name + ".pdf";

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
        private Task ErrorHandler(ITelegramBotClient arg1, Exception arg2, CancellationToken arg3)
        {
            throw new NotImplementedException();
        }
        public static string getReportSSRSPDFExcel(string reportname, string param, string filename, string pathreport)
        {
            try
            {
                string url = "http://10.15.2.60/ReportServerEOD?/" + pathreport + "/" + reportname + "&rs:Command=Render&rs:Format=EXCELOPENXML&rc:OutputFormat=XLS" + param;


                System.Net.HttpWebRequest Req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                Req.Credentials = new NetworkCredential("Administrator", "Jakarta01");
                Req.Method = "GET";

                string path = AppDomain.CurrentDomain.BaseDirectory + "report\\" + filename + ".xlsx";

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
        public static async Task<string> uploadGoogle(string path, string nameFile, string mimetype)
        {
            string uploadedFileId = "";
            try
            {
                var token = new FileDataStore("UserCredentialStoragePath", true);
                UserCredential credential;
                string[] scopes = new string[] { DriveService.Scope.Drive, DriveService.Scope.DriveFile, };
                await using (var stream = new FileStream(PathToCredentials, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                                    new ClientSecrets
                                    {
                                        ClientId = "134759934424-kgi3i9kkqt7g73fdak9tuqadj0ijeus1.apps.googleusercontent.com",
                                        ClientSecret = "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",
                                    },
                                    new[] { DriveService.Scope.Drive },
                                    "user",
                                CancellationToken.None
                                , new FileDataStore(AppDomain.CurrentDomain.BaseDirectory, false)).Result;
                }
                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "AwsomeAoo"
                });

                var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                {
                    Name = nameFile,
                    Parents = new List<string> { "1QujweILVyQ2hQAe5-LHZ2hftwTF9H1HJ" }
                };

                await using (var fssource = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    var request = service.Files.Create(fileMetadata, fssource, mimetype);
                    request.Fields = "*";
                    var result = await request.UploadAsync(CancellationToken.None);
                    if (result.Status == Google.Apis.Upload.UploadStatus.Failed)
                    {
                        Console.WriteLine($"Error uploading file: {result.Exception.Message}");
                        SendMessage("6289630870658@c.us", $"Error uploading file: {result.Exception.Message}");
                    }
                    uploadedFileId = request.ResponseBody?.Id;
                }

                //var request = service.Files.List();
                //request.Q = "parent in '1QujweILVyQ2hQAe5-LHZ2hftwTF9H1HJ' and mimeType = 'application/vnd.google-apps.document'";
                //var result = await request.ExecuteAsync();
                //foreach (var driveFile in result.Files)
                //{
                //    Console.WriteLine($"{driveFile.Name} {driveFile.MimeType} {driveFile.Id}");
                //}
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                SendMessage("6289630870658@c.us", $"Error uploading file: {ex.Message}");

            }
            return uploadedFileId;
        }
        public class data_csv
        {
            public string AccountNo { get; set; }
            public string Date { get; set; }
            public string ValDate { get; set; }
            public string TransactionCode { get; set; }
            public string Description1 { get; set; }
            public string Description2 { get; set; }
            public string ReferenceNo { get; set; }
            public string Debit { get; set; }
            public string Credit { get; set; }
            public string keterangan { get; set; }

        }
        private static void sendFileTelegram(string chatId, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;

            var client = new RestClient("https://api.telegram.org/bot5278461864:AAGwGEV3aJx8ZFKFfRaFB6TMc4mOem-uoc8/sendDocument");
            RestRequest requestWa = new RestRequest("https://api.telegram.org/bot5278461864:AAGwGEV3aJx8ZFKFfRaFB6TMc4mOem-uoc8/sendDocument", Method.Post);


            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "multipart/form-data");
            requestWa.AddParameter("chat_id", chatId);
            requestWa.AddFile("document", body);
            var responseWa = client.ExecutePostAsync(requestWa);
            Console.WriteLine(responseWa.Result.Content);
        }
        private static void SendFile(string chatId, string caption, string body)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/sendLink?token=jkdjtwjkwq2gfkac");

            RestRequest requestWa = new RestRequest("https://api.chat-api.com/instance127354/sendLink?token=jkdjtwjkwq2gfkac", Method.Post);

            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("body", body);
            requestWa.AddParameter("previewBase64", null);
            requestWa.AddParameter("title", caption);
            requestWa.AddParameter("chatId", chatId);
            var responseWa = client.ExecutePostAsync(requestWa);
            Console.WriteLine(responseWa.Result.Content);

        }
        public class MessageChat
        {
            public string id { get; set; }
            public string body { get; set; }
            public string fromMe { get; set; }
            public string self { get; set; }
            public string isForwarded { get; set; }
            public string author { get; set; }
            public double time { get; set; }
            public string chatId { get; set; }
            public int messageNumber { get; set; }
            public string type { get; set; }
            public string senderName { get; set; }
            public string caption { get; set; }
            public string quotedMsgBody { get; set; }
            public string quotedMsgId { get; set; }
            public string quotedMsgType { get; set; }
            public string chatName { get; set; }
        }
        public class ResponseChat
        {
            public IEnumerable<MessageChat> messages { get; set; }
            public int lastMessageNumber { get; set; }
        }
        private static string SendMessage(string chatId, string body)
        {
            var client = new RestClient("https://api.chat-api.com/instance127354/sendMessage?token=jkdjtwjkwq2gfkac");

            RestRequest requestWa = new RestRequest("https://api.chat-api.com/instance127354/sendMessage?token=jkdjtwjkwq2gfkac", Method.Post);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("chatId", chatId);
            requestWa.AddParameter("body", body);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }
        private static string GetMessageList(string chatId, int lastMessageNumber)
        {
            string url = "https://api.chat-api.com/instance127354/messages?token=jkdjtwjkwq2gfkac&lastMessageNumber=" + lastMessageNumber + "&chatId=" + chatId;
            var client = new RestClient(url);
            var requestWa = new RestRequest(url, Method.Get);
            var responseWa = client.ExecuteGetAsync(requestWa);
            return responseWa.Result.Content;

        }
        private static string imgBB(string data)
        {


            var client = new RestClient("https://api.imgbb.com/1/upload?key=9162df8da9fab25bfb9bc762ee774d10");

            RestRequest request = new RestRequest("https://api.imgbb.com/1/upload?key=9162df8da9fab25bfb9bc762ee774d10", Method.Post);
            request.Timeout = -1;
            request.AddFile("image", "E:/ServiceDariTimahDBMS/ServiceNotifSendEmail/" + data);
            request.AddParameter("expiration", "120");
            var response = client.ExecutePostAsync(request);
            return (response.Result.Content);
        }
        public class data_imgBB
        {
            public string url { get; set; }
        }
        public class ResponseImgBB
        {
            public data_imgBB data { get; set; }
        }
        public class data_bukti_transfer
        {
            public string AccountNo { get; set; }
            public string Date { get; set; }
            public string ValDate { get; set; }
            public string TransactionCode { get; set; }
            public string Description1 { get; set; }
            public string Description2 { get; set; }
            public string ReferenceNo { get; set; }
            public string Debit { get; set; }
            public string Credit { get; set; }

        }
        private static string monitoringServices(string servicename, string servicedescription, string servicelocation, string appstatus)
        {
            string jsonString = "{" +
                                "\"service_name\" : \"" + servicename + "\"," +
                                "\"service_description\": \"" + servicedescription + "\"," +
                                "\"service_location\":\"" + servicelocation + "\"," +
                                "\"app_status\":\"" + appstatus + "\"," +
                                "}";
            var client = new RestClient("http://10.10.10.99:84/api/ServiceStatus");

            RestRequest requestWa = new RestRequest("http://10.10.10.99:84/api/ServiceStatus", Method.Post);
            requestWa.Timeout = -1;
            requestWa.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            requestWa.AddParameter("data", jsonString);
            var responseWa = client.ExecutePostAsync(requestWa);
            return (responseWa.Result.Content);
        }

    }
}
