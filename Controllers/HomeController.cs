using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Webview_IRT.Models;
using System.Data.SqlClient;
using System.Text.Json;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;
using System.Net.Http;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Web.Http;
using System.Data;

namespace RTSM_OLSingleArm.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        public string connectionString;
        readonly IConfiguration _configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }


        public IActionResult Index()
        {

            if (HttpContext.Session.GetString("suserid") != null)
            {
                int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
                // string connection = "server=VMSQL2K16B.cro.com\\DEVDB;database=A_IRT_LIB_DEV;UID=appsid2;PWD=Kaz!Sal.01;Integrated Security=False";
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                // SQL query to retrieve data
                string sql = "SELECT SITEID, SUM(case when STATUS_INFO = 'Screened' then 1 else 0 end )as 'Total Screen Number', SUM(case when STATUS_INFO = 'Screen Failed' then 1 else 0 end )as 'Total Screen Fail Number', SUM(case when STATUS_INFO = 'RAND' OR STATUS_INFO = 'Randomized' then 1 else 0 end )as 'Total Randomization Number'  from BIL_SUBJ WHERE (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " Group by SITEID ";
                string sql2 = "SELECT COUNT(*) FROM BIL_SUBJ WHERE  SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " AND (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND (STATUS_INFO = 'RAND' OR STATUS_INFO = 'Randomized')";
                string sql3 = "SELECT PIDet, PIDesc FROM ProfileInfo WHERE  SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " AND PIType = 'StopAt'";
                // List to store the data
                List<Subject> reports = new List<Subject>();


                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    SqlCommand cmd = new SqlCommand(sql, con);
                    //SqlCommand cmd1 = new SqlCommand(sql2, con);
                    SqlCommand cmd3 = new SqlCommand(sql3, con);

                    con.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();
                    // SqlDataReader rdr3 = cmd3.ExecuteReader();

                    while (rdr.Read())
                    {
                        Subject report = new Subject
                        {
                            SITEID = rdr["SITEID"].ToString(),
                            TotalScreenNumber = (int)rdr["Total Screen Number"],
                            TotalScreenFailNumber = (int)(rdr["Total Screen Fail Number"]),
                            TotalRandomizationNumber = (int)(rdr["Total Randomization Number"])
                        };
                        reports.Add(report);

                    } rdr.Close();
                    SqlDataReader rdr3 = cmd3.ExecuteReader();
                    while (rdr3.Read())
                    {

                        {
                            ViewBag.TotalRandomization = rdr3["PIDet"];
                        }

                    }
                    rdr3.Close();
                    using (SqlCommand cmd1 = new SqlCommand(sql2, con))
                    {
                        ViewBag.TotalNumber = (int)cmd1.ExecuteScalar();
                    }

                    StudyName();
                    string SITEID = HttpContext.Session.GetString("sesCenter");
                    string userid = HttpContext.Session.GetString("suserid");
                    //string KITSET = ShipmentKey(SITEID);
                    bool emailSent = false;
                    //if (KITSET != "")
                    //{
                        if (HttpContext.Session.GetString("emailSent") != "true") // Check if email has not been sent
                        {
                                string KITSET = ShipmentKey(SITEID);
                                if (KITSET != "") { 
                                    string msgBody = "Protocol: Webview RTSM - Test" + Environment.NewLine;
                                    msgBody += "Site ID: " + SITEID + Environment.NewLine;
                                    msgBody += "This is to inform you that certain shipments in our Webview RTSM system have not been received for over five business days." + Environment.NewLine;
                                    msgBody += "Please log in to the system and update the status of these shipments.";
                                    string retSupp = "";
                                    string retSite = "";
                                    //retSupp = GetEmailByGrp(spkey, "(All)", "D");
                                    retSite = GetEmailByGrp(SPKEY, SITEID, "S");
                                    //if (retSupp == "")
                                    //{

                                    SendEmail(retSite + ";" + "jacobk@amarexcro.com", "Webview RTSM - Kit Shipment", msgBody);

                                    //}

                                    //else
                                    //{
                                    //    SendEmail(retSupp + ";" + "jacobk@amarexcro.com", "Webview RTSM - Kit Shipment - Site " + SITEID + " - Low Inv with Shipped status", msgBody);
                                    //}
                                    emailSent = true; // Set the flag to indicate that the email has been sent
                                    string[] Shipment = KITSET.Split(';');

                                    string insertsql = "INSERT INTO Log_Activity (SPKEY, USERID, LOG_DESC, LOG_TYPE, KITSET, SITEID) VALUES (@SPKEY, @USERID, @LOG_DESC,  @LOG_TYPE, @KITSET, @SITEID)";
                                    SqlConnection conn = new SqlConnection(connectionString);
                                    // SqlCommand cmd = new SqlCommand(sql, conn);
                                    using (conn)
                                    {
                                        conn.Open();
                                        for (int i = 0; i < Shipment.Count(); i++)
                                        {
                                            SqlCommand insertcmd = new SqlCommand(insertsql, conn);
                                            insertcmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                                            insertcmd.Parameters.AddWithValue("@USERID", userid );
                                            insertcmd.Parameters.AddWithValue("@LOG_DESC", "Email Sent");
                                            insertcmd.Parameters.AddWithValue("@LOG_TYPE", "Email-Notification");
                                            insertcmd.Parameters.AddWithValue("@KITSET", Shipment[i]);
                                            insertcmd.Parameters.AddWithValue("@SITEID", SITEID);
                                            insertcmd.ExecuteNonQuery();
                                        }
                                    }

                                    HttpContext.Session.SetString("emailSent", "true"); // Store the flag in session
                                    TempData["Modal"] = "Open";

                                }
                        }
                   //}
                }

                return View(reports);
            }
            else
                return View();
        }
        public IActionResult CloseButton()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string KITSET = ShipmentKey(SITEID);
            string[] Shipment = KITSET.Split(';');

            string sql = "INSERT INTO Log_Activity (SPKEY, USERID, LOG_DESC, LOG_TYPE, KITSET, SITEID) VALUES (@SPKEY, @USERID, @LOG_DESC,  @LOG_TYPE, @KITSET, @SITEID)";
            SqlConnection conn = new SqlConnection(connectionString);
            // SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                for (int i = 0; i < Shipment.Count(); i++)
                {
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                    cmd.Parameters.AddWithValue("@USERID", userid);
                    cmd.Parameters.AddWithValue("@LOG_DESC", "Closed Button Clicked");
                    cmd.Parameters.AddWithValue("@LOG_TYPE", "Popup-Notification");
                    cmd.Parameters.AddWithValue("@KITSET", Shipment[i]);
                    cmd.Parameters.AddWithValue("@SITEID", SITEID);
                    cmd.ExecuteNonQuery();
                }
            }
            return RedirectToAction("Index");

        }
        [HttpPost]
        public IActionResult ProceedButton()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string KITSET = ShipmentKey(SITEID);
            string[] Shipment = KITSET.Split(';');

            string sql = "INSERT INTO Log_Activity (SPKEY, USERID, LOG_DESC,  LOG_TYPE, KITSET, SITEID) VALUES (@SPKEY, @USERID, @LOG_DESC,  @LOG_TYPE, @KITSET, @SITEID)";
            SqlConnection conn = new SqlConnection(connectionString);
           
            using (conn)
            {
                conn.Open();
                for (int i = 0; i < Shipment.Count(); i++)
                {
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                    cmd.Parameters.AddWithValue("@USERID", userid);
                    cmd.Parameters.AddWithValue("@LOG_DESC", "Proceed Button Clicked");
                    cmd.Parameters.AddWithValue("@LOG_TYPE", "Popup-Notification");
                    cmd.Parameters.AddWithValue("@KITSET", Shipment[i]);
                    cmd.Parameters.AddWithValue("@SITEID", SITEID);
                    cmd.ExecuteNonQuery();
                }
            }
            return RedirectToAction("ReceiveIPTabs", "SupplyManagement");

        }


        public string ShipmentKey(string SITEID)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection con = new SqlConnection(connectionString);
            string sqlState1 = "SELECT *, DATEDIFF(dd,BIL_IP_RANGE.SENDDATE,getdate()) AS SENDDAYS FROM BIL_IP_RANGE WHERE (SITEID = '" + SITEID + "') AND (KITSTAT = 'Shipped')";
            SqlCommand cmd2 = new SqlCommand(sqlState1, con);
            
            bool emailSent = false;
            int count = 0;
            string KITSET = "";
            using (con)
            {
                con.Open();
                SqlDataReader reader = cmd2.ExecuteReader();
                while (reader.Read())
                {
                    if (Convert.ToInt16(reader["SENDDAYS"]) > 5)
                    {
                        if(KITSET == "")
                          KITSET = reader["KITSET"].ToString();
                        else
                            KITSET += ";"+reader["KITSET"].ToString();
                        count++;

                    }
                }
                con.Close();
            }
            return KITSET;
            }
        public string GetEmailByGrp(int SPKEY, string SITEID, string type)
        {
            SqlConnection con = new SqlConnection(_configuration.GetConnectionString("amarexDbConnStr"));
            string rtnVal = "";
            var sql = "";
            sql = "SELECT zSecurityID.User_Email FROM zSecurityID INNER JOIN zSecUserIDCenter ";
            sql += "ON zSecurityID.UserID = dbo.zSecUserIDCenter.UserID WHERE (zSecUserIDCenter.SPKEY = " + SPKEY;
            sql += ") AND (Center_Num = '" + SITEID + "')  AND (dbo.zSecUserIDCenter.Center_Lvl = '" + type + "')";
            SqlCommand cmd = new SqlCommand(sql, con);
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rtnVal == "")
                    {
                        rtnVal = rdr["User_Email"].ToString();
                    }
                    else
                    {
                        rtnVal += ";" + rdr["User_Email"].ToString();
                    }
                }
            }
            con.Close();

            return rtnVal;
        }
        public void SendEmail(string SendTo, string Subject, string message)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);

            SqlCommand EmailCmd = new SqlCommand("NotifySend2", conn);
            EmailCmd.CommandType = CommandType.StoredProcedure;
            EmailCmd.Parameters.Add("@SENDMAILTO", SqlDbType.VarChar).Value = SendTo;
            EmailCmd.Parameters.Add("@SUBJ", SqlDbType.VarChar).Value = Subject;
            EmailCmd.Parameters.Add("@MSGBODY", SqlDbType.VarChar).Value = message;


            using (conn)
            {
                conn.Open();
                EmailCmd.ExecuteReader();
            }

        }
        public void StudyName()
        {
            connectionString = _configuration.GetConnectionString("AmarexDbConnStr");
            string sql2 = "SELECT StudyName FROM STUDY_PROFILE2 WHERE  SPKEY = " + HttpContext.Session.GetString("sesSPKey") + "";
            SqlConnection con = new SqlConnection(connectionString);
            using (con)
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(sql2, con);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    String study = rdr["StudyName"].ToString();
                    TempData["StudyName"] = rdr["StudyName"].ToString();
                }

            }
            //Tempdata.peak["StudyName]

            
        }
        
        public IActionResult SignOut()
        {
            Logout();
            //HttpContext.Session.Set("suserid", null);
            HttpContext.Session.Clear();
            return Redirect("https://webviewsso.com:8443/WebViewSSO_UAT/");
        }

        [HttpPost]
        public async Task Logout()
        {
            //await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme); // still dont know why redirect uri is not working for binging the user back to the home page after logging out
            //return Index(); // aparently this is a problem
            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme, new AuthenticationProperties { RedirectUri = "/" });

            // return;
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult ResetSession()
        {
             // Clear all session values

            return Json(new { success = true });
        }

        public IActionResult MyNotifications()
        {
            List<Email_Notifications> notifList = new List<Email_Notifications>();
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string sql = "SELECT KEYID, PIDet, PIDESC, PIType, UserID, Enable FROM Email_Notifications WHERE UserID = '" + HttpContext.Session.GetString("suserid") + "' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
           // connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
           
                SqlCommand cmd = new SqlCommand(sql, conn);
                using (conn)
                {
                    conn.Open();
                   
                   SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                    Email_Notifications temp = new Email_Notifications();
                        temp.KEYID = (int)rdr["KEYID"];
                        temp.PIDet = rdr["PIDet"].ToString();
                        temp.PIType = rdr["PIType"].ToString();
                        temp.PIDESC = rdr["PIDESC"].ToString();
                    object enableValue = rdr["Enable"];
                    if (enableValue != DBNull.Value)
                    {
                        temp.Enable = (bool)enableValue;
                    }

                    notifList.Add(temp);
                    }
                }
                return View(notifList);
            }
        [HttpPost]
        public IActionResult UpdateMyNotification(int KEYID,  bool Enable)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql = "UPDATE Email_Notifications SET [Enable] = '" + Enable + "', CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER WHERE (KEYID = " + KEYID + ")";
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@CHANGEUSER", HttpContext.Session.GetString("suserid"));
                cmd.ExecuteNonQuery();
                string Message = "Notification updated";
                TempData["Message"] = Message;
            }
            return RedirectToAction("MyNotifications");
        }
    }
}
