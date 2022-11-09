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
using Microsoft.Extensions.Configuration;
using System.Net.Mail;

namespace WorkerEmail
{
    public class Worker : BackgroundService
    {
        //tutorial google drive
        //https://www.youtube.com/watch?v=pHOweM1Gl6c
        //create project
        //open api drive

        private readonly ILogger<Worker> _logger;
        private const string PathToCredentials = @"D:\ServiceSendEmailMonex\credentials.json";

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

        public IConfiguration Configuration { get; }
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    var hour = DateTime.Now.Hour;
                    var minute = DateTime.Now.Minute;
                    var day = DateTime.Now.DayOfWeek;
                    sendEmail();
                    if (day.ToString().ToLower() != "sunday")
                    {

                        if (hour == 8 && minute == 01)
                        {
                            monitoringServices("kbi_sendemailmonex", "Service untuk upload file monex ke google drive dan otomatis kirim email", "10.10.10.79", "Live");

                            var connectionString = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ConnectionStrings")["SKDConnection"];
                            //GET BUSINES DATE
                            DateTime businessdate = DateTime.Now;
                            string revision = "";

                            var querry = "SELECT DateValue FROM SKD.Parameter WHERE Code = 'LastEOD'";

                            using (SqlConnection connection = new SqlConnection(connectionString))
                            {
                                SqlCommand command = new SqlCommand(querry, connection);
                                connection.Open();
                                command.CommandTimeout = 1800;
                                SqlDataReader reader = command.ExecuteReader();
                                while (reader.Read())
                                {
                                    businessdate = Convert.ToDateTime((reader["DateValue"]));
                                }
                                reader.Close();
                                connection.Dispose();
                                connection.Close();
                            }
                            _logger.LogInformation(businessdate.ToString());
                            var querryrevision = "SELECT EODRevision FROM SKD.EODRevision WHERE BusinessDate = '" + businessdate.ToString("yyyy-MM-dd") + "'";
                            using (SqlConnection connection = new SqlConnection(connectionString))
                            {
                                SqlCommand command = new SqlCommand(querryrevision, connection);
                                connection.Open();
                                command.CommandTimeout = 1800;
                                SqlDataReader reader = command.ExecuteReader();
                                while (reader.Read())
                                {
                                    revision = (reader["EODRevision"]).ToString();
                                }
                                reader.Close();
                                connection.Dispose();
                                connection.Close();
                            }
                            //END
                            _logger.LogInformation(revision.ToString());
                            var shortdate = businessdate.ToString("d").Split('/');
                            var containname = businessdate.ToString("yyyyMMdd") + "-" + revision + "_052";
                            //string reaultemail = sendEmail();

                            string[] files = Directory.GetFiles("D:\\Share\\EODReports\\AK\\" + businessdate.ToString("yyyy") + "\\" + shortdate[0] + "\\" + shortdate[1], "*.pdf");
                            foreach (var file in files)
                            {
                                if (file.Contains(containname))
                                {
                                    string[] filenames = file.ToString().Split("\\");
                                    string filename = filenames[filenames.Length - 1];

                                    var returnpath = await uploadGoogle(file, filename, "application/x-pdf");
                                    _logger.LogInformation(returnpath);
                                }
                            }
                            sendEmail();
                            _logger.LogInformation("success send");
                        }
                        else if (hour == 10 && minute == 01)
                        {
                            monitoringServices("kbi_sendemailmonex", "Service untuk delete file monex di google drive", "10.10.10.79", "Live");

                            await deletefiles();
                            _logger.LogInformation("success delete");
                        }

                    }
                }
                catch (Exception ex)
                {
                    _logger.LogInformation(ex.Message);
                }
                await Task.Delay(60000, stoppingToken);
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
                    Parents = new List<string> { "1ur44Pj6kw0luk6GFpiHZTcw_FJnCx_Rd" }
                };

                await using (var fssource = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    var request = service.Files.Create(fileMetadata, fssource, mimetype);
                    request.Fields = "*";
                    var result = await request.UploadAsync(CancellationToken.None);
                    if (result.Status == Google.Apis.Upload.UploadStatus.Failed)
                    {
                        Console.WriteLine($"Error uploading file: {result.Exception.Message}");
                    }
                    uploadedFileId = request.ResponseBody?.Id;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return uploadedFileId;
        }
        public static async Task<string> deletefiles()
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

                var x = service.Files.List();
                x.Q = "parents in '1ur44Pj6kw0luk6GFpiHZTcw_FJnCx_Rd' and mimeType = 'application/x-pdf'";
                var result = x.ExecuteAsync();
                //service.Files.Delete("1yZQxuyJIfsUlMQZnV3OhIqyYa4JTOtgD").Execute();
                //FilesResource.DeleteRequest request1 = service.Files.Delete(id);

                //var request = service.Files.List();
                //request.Q = "parents in '1ur44Pj6kw0luk6GFpiHZTcw_FJnCx_Rd' mimeType = 'application/vnd.google-apps.folder'";
                //var result = await request.ExecuteAsync();
                for (int i = 0; i < result.Result.Files.Count(); i++)
                {
                    service.Files.Delete(result.Result.Files[i].Id).Execute();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }
            return "ok";
        }
        public static string sendEmail()
        {
            try
            {
                //send email
                var email1 = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ParameterEmail")["email1"];
                var email2 = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ParameterEmail")["email2"];
                var email3 = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ParameterEmail")["email3"];
                var email4 = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ParameterEmail")["email4"];

                //MailAddress addressFrom = "pb@ptkbi.com";

                //MailAddress[] addressTo = new MailAddress[3];
                //addressTo[0] = email1;
                //addressTo[1] = email2;
                //addressTo[2] = email3;
                ////addressTo[3] = email4;


                //MailMessage message1 = new MailMessage(addressFrom, addressTo);
                ////message1.Cc = "jujuk1020@gmail.com";
                ////message1.Cc = "drp@ptkbi.com";
                //message1.Subject = "Pengiriman report AK PT Monex Investindo Futures";
                //message1.BodyText = "\r\nSelamat Pagi Bapak / Ibu,\r\n" + "Sehubungan dengan maksimal jumlah size dalam pengiriman email kami, laporan harian kami kirimkan dalam bentuk Link Google Drive. untuk dapat membuka laporan harian, bapak / ibu dapat mengunduh file pdf pada link terlampir.\r\nhttps://drive.google.com/drive/folders/1ur44Pj6kw0luk6GFpiHZTcw_FJnCx_Rd?usp=sharing\r\nFile ini akan otomatis terhapus setelah jam 10.00\n\r***Pesan ini dikirim secara otomatis***\r\n" + "Terima kasih\r\nPT.KBI";


                //SmtpClient smtp = new SmtpClient();
                //smtp.Host = "10.10.10.2";
                //smtp.ConnectionProtocols = Spire.Email.IMap.ConnectionProtocols.None;
                //smtp.Username = addressFrom.Address;
                ////smtp.Password = "Jakarta2021";
                //smtp.Port = 25;
                //smtp.SendOne(message1);
                //return ("From : " + message1.From.ToString() + "\nTo: " + message1.To.ToString() + "\nSubject: " + message1.Subject + "\n*** BODY ***\n" + message1.BodyText + "\nMessage Sent.");
                string file = AppDomain.CurrentDomain.BaseDirectory + "\\index.html";
                string text = File.ReadAllText(file);

                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                message.From = new MailAddress("pb@ptkbi.com");
                message.To.Add(new MailAddress(email1));
                message.Bcc.Add(new MailAddress(email2));
                message.Bcc.Add(new MailAddress(email4));
                message.Subject = "Pengiriman report AK PT Monex Investindo Futures";
                message.IsBodyHtml = true; //to make message body as html  
                message.Body = text;
                smtp.Port = 25;
                smtp.Host = "10.10.10.2"; //for gmail host  
                smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                //smtp.Credentials = new NetworkCredential("automatic_ptkbi@outlook.com", "Jakarta2021");
                smtp.EnableSsl = false;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
                return ("oke");
            }
            catch (Exception ex)
            {
                return (ex.Message);
            }

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
