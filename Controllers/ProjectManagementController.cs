using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Threading.Tasks;
using Webview_IRT.Models;

namespace RTSM_OLSingleArm.Controllers
{
    public class ProjectManagementController : Controller
    {
        public string connectionString;
        readonly IConfiguration _configuration;
        public ProjectManagementController(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        
        public IActionResult Test()
        {
            return View();
        }
        
        public IActionResult ProjectManagementHome()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql = "SELECT ROW_KEY, SITEID, SUBJID, BRTHDTC, SEX, ICDTC,  STATUS_INFO, PMCOMDATE, ARM FROM BIL_SUBJ WHERE STATUS_INFO = 'Randomized' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID  ";
            string sql2 = "SELECT ROW_KEY, SITEID, SUBJID, BRTHDTC, SEX, ICDTC,  STATUS_INFO, PMCOMDATE FROM BIL_SUBJ WHERE (STATUS_INFO = 'Screened' OR STATUS_INFO = 'Screen Failed') AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID   ";
                //String statsql = "SELECT STATUS_INFO, BRTHDTC, SEX, SCRNDTC, SUBJID, SITEID, ICDTC FROM BIL_SUBJ WHERE STATUS_INFO = 'Screen' ";
                SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand(sql, con);
                SqlCommand cmd2 = new SqlCommand(sql2, con);

                var list = new ScrnList();
                var randList = new List<SubjStat>();
                var stat = new List<Scrnstat>();
                

                using (con)
                {
                    con.Open();
                //cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
                //cmd.Parameters.AddWithValue("@SITEID", HttpContext.Session.GetString("sesCenter")); 
                SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var subjects = new SubjStat();
                        subjects.ROW_KEY = (int)rdr["ROW_KEY"];
                        subjects.SUBJID = rdr["SUBJID"].ToString();
                        subjects.BRTHDTC = rdr["BRTHDTC"].ToString();
                        subjects.SEX = rdr["SEX"].ToString();
                        subjects.ICDTC = rdr["ICDTC"].ToString();
                        subjects.ARM = rdr["ARM"].ToString();
                        subjects.STATUS_INFO = rdr["STATUS_INFO"].ToString();
                        if (!Convert.IsDBNull(rdr["PMCOMDATE"]) && rdr["PMCOMDATE"] != null)
                        {
                            subjects.PMCOMDATE = (DateTime)rdr["PMCOMDATE"];
                        }
                        
                    randList.Add(subjects);
                    }
                    rdr.Close();
               // cmd2.Parameters.AddWithValue("@SPKEY", SPKEY);
                //cmd2.Parameters.AddWithValue("@SITEID", SITEID);
                SqlDataReader rdr3 = cmd2.ExecuteReader();
                    while (rdr3.Read())
                    {
                        var subjects = new Scrnstat();
                        subjects.ROW_KEY = (int)rdr3["ROW_KEY"];
                        subjects.SUBJID = rdr3["SUBJID"].ToString();
                        subjects.BRTHDTC = rdr3["BRTHDTC"].ToString();
                        subjects.SEX = rdr3["SEX"].ToString();
                        subjects.ICDTC = rdr3["ICDTC"].ToString();
                       // subjects.SCRNDTC = rdr3["SCRNDTC"].ToString();
                        subjects.STATUS_INFO = rdr3["STATUS_INFO"].ToString();
                    if (!Convert.IsDBNull(rdr3["PMCOMDATE"]) && rdr3["PMCOMDATE"] != null)
                    {
                        subjects.PMCOMDATE = (DateTime)rdr3["PMCOMDATE"];
                    }

                    stat.Add(subjects);
                    }
                    rdr3.Close();

                }
                con.Close();
                list.subjList = randList;
                list.statList = stat;

                return View(list);
        }
        public Subject getScreen(int ROW_KEY)
        {
            Subject tempSub = new Subject();
            tempSub.ROW_KEY = ROW_KEY;
            string sql = "SELECT ARM, ROW_KEY, SITEID, SUBJID, BRTHDTC, SEX, ICDTC, STATUS_INFO, PM_COMM, SFDATE FROM BIL_SUBJ where ROW_KEY = @ROW_KEY";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@ROW_KEY", ROW_KEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    tempSub.ROW_KEY = (int)rdr["ROW_KEY"];
                    tempSub.SITEID = rdr["SITEID"].ToString();
                    tempSub.SUBJID = rdr["SUBJID"].ToString();
                    tempSub.BRTHDTC = rdr["BRTHDTC"].ToString();
                    tempSub.SEX = rdr["SEX"].ToString();
                    tempSub.ICDTC = rdr["ICDTC"].ToString();
                   // tempSub.SCRNDTC = rdr["SCRNDTC"].ToString();
                    tempSub.STATUS_INFO = rdr["STATUS_INFO"].ToString();
                    tempSub.PM_COMM = rdr["PM_COMM"].ToString();
                    tempSub.SFDATE = rdr["SFDATE"].ToString();
                    tempSub.ARM = rdr["ARM"].ToString();

                    // tempDev.status = rdr["STATUS"].ToString();
                }
                rdr.Close();
            }
            return tempSub;
        }
        [HttpGet]
        public IActionResult EditScreen(int ROW_KEY)
        {
            // StudyDropdown();
            return View(getScreen(ROW_KEY));
        }

        
        [HttpPost]
        public IActionResult EditScreen(string SITEID, int ROW_KEY, string SUBJID, string BRTHDTC, string SEX, string ICDTC,  string  STATUS_INFO, string PM_COMM, string SFDATE, string username, string password)
        {
            string userid = HttpContext.Session.GetString("suserid");
            if (IsValidUser(username, password) && (string.Equals(userid, username, StringComparison.OrdinalIgnoreCase) || userid == "sidran"))
            {
                int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
                
                //SITEID = HttpContext.Session.GetString("sesCenter");
                string sql = "Update BIL_SUBJ set SUBJID = @SUBJID, BRTHDTC = @BRTHDTC, SEX = @SEX, ICDTC = @ICDTC, STATUS_INFO = @STATUS_INFO, PM_COMM = @PM_COMM, CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER, SFDATE = @SFDATE, PMCOMDATE = SYSDATETIME() where ROW_KEY = @ROW_KEY";
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand(sql, conn);
                using (conn)
                {
                    conn.Open();

                    cmd.Parameters.AddWithValue("@SUBJID", SUBJID);
                    cmd.Parameters.AddWithValue("@BRTHDTC", BRTHDTC);
                    cmd.Parameters.AddWithValue("@SEX", SEX);
                    cmd.Parameters.AddWithValue("@ICDTC", (object)ICDTC ?? DBNull.Value);

                    //cmd.Parameters.AddWithValue("@SCRNDTC", SCRNDTC);
                    cmd.Parameters.AddWithValue("@ROW_KEY", ROW_KEY);
                    cmd.Parameters.AddWithValue("@STATUS_INFO", STATUS_INFO);
                    cmd.Parameters.AddWithValue("@PM_COMM", PM_COMM);
                    cmd.Parameters.AddWithValue("@CHANGEUSER", userid);
                    object sfdate = (object)SFDATE ?? DBNull.Value;
                    cmd.Parameters.AddWithValue("@SFDATE", sfdate);



                    cmd.ExecuteNonQuery();
                }
                TempData["Message"] = "Subject " + SUBJID + " Information updated";
                string message = "Protocol: Webview RTSM - Test"  + Environment.NewLine;
                message += "The Subject information has been successfully updated  for the assigned Subject ID " + SUBJID + ".";
                string subject = "Webview RTSM - Test For [Amarex][Webview RTSM - Test] - Subject " + SUBJID +"";
                SendEmail(GetUserEmail(), subject, message);
                return RedirectToAction("ProjectManagementHome");
            }
            else
            {

                // Invalid username or password
                string errorMessage = "Invalid username or password.";
                //ViewBag.ErrorMessage = errorMessage;
                TempData["ErrorMessage"] = errorMessage;
                return View();
            }

        }


        [HttpGet]
        public IActionResult EditRand(int ROW_KEY)
        {
            // StudyDropdown();
            return View(getScreen(ROW_KEY));
        }


        [HttpPost]
        public IActionResult EditRand(string SITEID, int ROW_KEY, string SUBJID, string BRTHDTC, string SEX, string ICDTC,  string STATUS_INFO, string PM_COMM, string username, string password)
        {
            string userid = HttpContext.Session.GetString("suserid");
            if (IsValidUser(username, password) && (string.Equals(userid, username, StringComparison.OrdinalIgnoreCase) || userid == "sidran"))
            {
                int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
                
                //string SITEID = HttpContext.Session.GetString("sesCenter");
                string sql = "Update BIL_SUBJ set SUBJID = @SUBJID, BRTHDTC = @BRTHDTC, SEX = @SEX, ICDTC = @ICDTC, STATUS_INFO = @STATUS_INFO, PM_COMM = @PM_COMM, CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER, PMCOMDATE = SYSDATETIME() where ROW_KEY = @ROW_KEY";
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand(sql, conn);
                using (conn)
                {
                    conn.Open();

                    cmd.Parameters.AddWithValue("@SUBJID", SUBJID);
                    cmd.Parameters.AddWithValue("@BRTHDTC", BRTHDTC);
                    cmd.Parameters.AddWithValue("@SEX", SEX);
                    cmd.Parameters.AddWithValue("@ICDTC", (object)ICDTC ?? DBNull.Value);
                   // cmd.Parameters.AddWithValue("@SCRNDTC", SCRNDTC);
                    cmd.Parameters.AddWithValue("@ROW_KEY", ROW_KEY);
                    cmd.Parameters.AddWithValue("@STATUS_INFO", STATUS_INFO);
                    cmd.Parameters.AddWithValue("@PM_COMM", PM_COMM);
                    cmd.Parameters.AddWithValue("@CHANGEUSER", userid);

                    cmd.ExecuteNonQuery();
                }
                TempData["Message"] = "Subject Info updated";
                string message = "Protocol: Webview RTSM - Test" + Environment.NewLine;
                message += "The Subject information has been successfully updated  for the assigned Subject ID " + SUBJID + ".";
                string subject = "Webview RTSM - Test For [Amarex][Webview RTSM - Test] - Subject " + SUBJID + "";
                SendEmail(GetUserEmail(), subject, message);
                return RedirectToAction("ProjectManagementHome");
            }
            else
            {
                // Invalid username or password
                string errorMessage = "Invalid username or password.";
                //ViewBag.ErrorMessage = errorMessage;
                TempData["ErrorMessage"] = errorMessage;
                return View();
            }
        }



        private DataTable CreateDataTable(string cmdText, string connection)
        {
            System.Data.DataTable dt = new DataTable();
            System.Data.SqlClient.SqlDataAdapter da = new SqlDataAdapter(cmdText, connection);
            da.Fill(dt);
            return dt;
        }


        public IActionResult SubjectInfo(object sender, EventArgs e)
        {
            XLWorkbook wb = new XLWorkbook();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DataTable dt = CreateDataTable("SELECT * FROM BIL_SUBJ Where STATUS_INFO = 'Screened' or STATUS_INFO = 'Screen Failed'", connectionString);
            wb.Worksheets.Add(dt, "SubjectInfo").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "SubjectInfo" + DateTime.Now.Year + ".xlsx");
            }
        }

        public IActionResult RandSubject(object sender, EventArgs e)
        {
            XLWorkbook wb = new XLWorkbook();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DataTable dt = CreateDataTable("SELECT * FROM BIL_SUBJ where STATUS_INFO = 'Randomized'", connectionString);
            wb.Worksheets.Add(dt, "RandSubject").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "RandSubject" + DateTime.Now.Year + ".xlsx");
            }
        }

        public void SendEmail(string SendTo, string Subject, string message)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);

            SqlCommand EmailCmd = new SqlCommand("NotifySend2", conn);
            EmailCmd.CommandType = CommandType.StoredProcedure;

            //    cmd.Parameters.Add("@pLogin", SqlDbType.VarChar).Value = User.FindFirst(ClaimTypes.NameIdentifier).Value;
            EmailCmd.Parameters.Add("@SENDMAILTO", SqlDbType.VarChar).Value = SendTo;
            EmailCmd.Parameters.Add("@SUBJ", SqlDbType.VarChar).Value = Subject;
            EmailCmd.Parameters.Add("@MSGBODY", SqlDbType.VarChar).Value = message;


            using (conn)
            {
                conn.Open();
                EmailCmd.ExecuteReader();
            }

        }
        public IActionResult ProblemReply()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            string sql = "SELECT PRBLMKEY, SITEID, SUBJID, PROBLEM_DESC, PM_COMM, EMAIL_SENT, ADDUSER, FORMAT(ADDDATE, 'MMM/dd/yyyy') AS DateReport, PROBLEM_STATUS  FROM BIL_PRBLM  WHERE SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            var model = new List<Reportproblem>();

            using (conn)
            {
                conn.Open();
                //cmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                //cmd.Parameters.AddWithValue("@SITEID", SITEID);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var subjects = new Reportproblem();
                    subjects.PRBLMKEY = (int)rdr["PRBLMKEY"];
                    subjects.SITEID = rdr["SITEID"].ToString();
                    subjects.SUBJID = rdr["SUBJID"].ToString();
                    subjects.PROBLEM_DESC = rdr["PROBLEM_DESC"].ToString();
                    subjects.PM_COMM = rdr["PM_COMM"].ToString();
                    subjects.EMAIL_SENT = rdr["EMAIL_SENT"].ToString();
                    //subjects.ADDDATE = (DateTime)rdr["ADDDATE"];
                    subjects.ADDUSER = rdr["ADDUSER"].ToString();
                    subjects.DateReport = rdr["DateReport"].ToString();
                    subjects.PROBLEM_STATUS = rdr["PROBLEM_STATUS"].ToString();
                    model.Add(subjects);
                }

            }
            conn.Close();
            return View(model);

        }
        public Reportproblem getReportinfo(int PRBLMKEY)
        {
            Reportproblem tempRep = new Reportproblem();
            tempRep.PRBLMKEY = PRBLMKEY;
            string sql = "SELECT PRBLMKEY, SITEID, SUBJID, PROBLEM_DESC, PM_COMM, EMAIL_SENT, ADDUSER, ADDDATE, PROBLEM_STATUS  FROM BIL_PRBLM where PRBLMKEY = @PRBLMKEY";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@PRBLMKEY", PRBLMKEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    tempRep.PRBLMKEY = (int)rdr["PRBLMKEY"];
                    tempRep.SITEID = rdr["SITEID"].ToString();
                    tempRep.SUBJID = rdr["SUBJID"].ToString();
                    tempRep.ADDUSER = rdr["ADDUSER"].ToString();
                    tempRep.ADDDATE = (DateTime)rdr["ADDDATE"];
                    tempRep.PROBLEM_DESC = rdr["PROBLEM_DESC"].ToString();
                    tempRep.EMAIL_SENT = rdr["EMAIL_SENT"].ToString();
                    tempRep.PM_COMM = rdr["PM_COMM"].ToString();
                    tempRep.PROBLEM_STATUS = rdr["PROBLEM_STATUS"].ToString();
                }
                rdr.Close();


            }
            return tempRep;

        }

        [HttpGet]
        public IActionResult ProblemResponse(int PRBLMKEY)
        {
            // StudyDropdown();
            return View(getReportinfo(PRBLMKEY));
        }


        [HttpPost]
        public IActionResult ProblemResponse(int PRBLMKEY, string SUBJID, string SITEID, string PROBLEM_DESC, string PM_COMM, string ADDUSER, string PROBLEM_STATUS, DateTime ADDDATE)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            
            string sql = "Update BIL_PRBLM set PM_COMM = @PM_COMM, CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER, EMAIL_SENT = @EMAIL_SENT, PROBLEM_STATUS = @PROBLEM_STATUS where PRBLMKEY = @PRBLMKEY";
            string requser = "SELECT ADDDUSER from BIL_PRBLM WHERE PRBLMKEY = @PRBLMKEY";

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();

                
                cmd.Parameters.AddWithValue("@PRBLMKEY", PRBLMKEY);
                cmd.Parameters.AddWithValue("@PROBLEM_STATUS", PROBLEM_STATUS);
                cmd.Parameters.AddWithValue("@PM_COMM", PM_COMM);
                cmd.Parameters.AddWithValue("@CHANGEUSER", userid);
                DateTime currentDateTime = DateTime.Now;
                cmd.Parameters.AddWithValue("@EMAIL_SENT", currentDateTime);


                cmd.ExecuteNonQuery();
            }
            string requestorID = GetRequesterEmail(PRBLMKEY);
            TempData["Message"] = "Email has been sent" ;
            string message = "Protocol: Webview RTSM - Test" + "\n";
            message += "Subject ID: " + SUBJID + "\n" + "Date Reported: " + ADDDATE + "\n" + "Problem Description: " + PROBLEM_DESC +"\n" + "\n" + "\n" + "Project Manager Comments: " + PM_COMM;
            string subject = "Webview RTSM - Reply to Problem Report Submitted for [Amarex][WebView RTSM -Test] Subject " + SUBJID;
            string email = GetEmail(userid) + ";" + GetEmail(requestorID);
            SendEmail(email, subject, message);
           // SendEmail(GetEmail(requestorID), subject,  message);
            return RedirectToAction("ProblemReply");
        }
        public IActionResult ReportProblemExcel(object sender, EventArgs e)
        {
            XLWorkbook wb = new XLWorkbook();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DataTable dt = CreateDataTable("SELECT * FROM BIL_PRBLM ", connectionString);
            wb.Worksheets.Add(dt, "ReportProblemExcel").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "ReportProblemExcel" + DateTime.Now.Year + ".xlsx");
            }
        }

        public IActionResult RequestHome()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            String sql = "SELECT PIKEY, PIDesc, PIDet, PIType FROM ProfileInfo WHERE PIType = 'RandStat' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            String sql2 = "SELECT PIKEY, PIDesc, PIDet, PIType FROM ProfileInfo WHERE PIType = 'ScreenStat' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            String sql3 = "SELECT PIKEY, PIDesc, PIDet, PIType FROM ProfileInfo WHERE PIType = 'StopAt' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            String sql4 = "SELECT PIKEY, PIDesc, PIDet, PIType  FROM ProfileInfo WHERE PIType = 'SiteScreen' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            String sql5 = "SELECT PIKEY, PIDesc, PIDet, PIType  FROM ProfileInfo WHERE PIType = 'SiteRand' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            String sql6 = "SELECT PIKEY, PIDesc, PIDet, PIType FROM ProfileInfo WHERE PIType = 'StopScreenAt' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlCommand cmd2 = new SqlCommand(sql2, con);
            SqlCommand cmd3 = new SqlCommand(sql3, con);
            SqlCommand cmd6 = new SqlCommand(sql6, con);

            var list = new PMmodel();
            var randList = new List<Randstat>();
            var screenstat = new List<ScreenInfo>();
            var stopat = new List<ProfileInfo>();
            var sitescr = new List<SiteScreenStatus>();
            var siternd = new List<SiteRandStatus>();
            var stopscrn = new List<StopScreenAt>();


            using (con)
            {
                con.Open();
                //cmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var subjects = new Randstat();
                    subjects.PIKEY = (int)rdr["PIKEY"];
                    subjects.PIDesc = rdr["PIDesc"].ToString();
                    subjects.PIDet = rdr["PIDet"].ToString();
                    subjects.PIType = rdr["PIType"].ToString();

                    if (rdr["PIDesc"].ToString().Equals("RandStat"))
                    {
                        subjects.PIDesc = "Randomization";
                    }
          
                    randList.Add(subjects);
                }
                rdr.Close();
                //cmd2.Parameters.AddWithValue("@SPKEY", SPKEY);
                
                SqlDataReader rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    var subjects = new ScreenInfo();
                    subjects.PIKEY = (int)rdr2["PIKEY"];
                    subjects.PIDesc = rdr2["PIDesc"].ToString();
                    subjects.PIDet = rdr2["PIDet"].ToString();
                    subjects.PIType = rdr2["PIType"].ToString();
                    if (rdr2["PIDesc"].ToString().Equals("ScreenStat"))
                    {
                        subjects.PIDesc = "Screening";
                    }

                    screenstat.Add(subjects);
                }
                rdr2.Close();
                //cmd3.Parameters.AddWithValue("@SPKEY", SPKEY);
                
                SqlDataReader rdr3 = cmd3.ExecuteReader();
                while (rdr3.Read())
                {
                    var subjects = new ProfileInfo();
                    subjects.PIKEY = (int)rdr3["PIKEY"];
                    subjects.PIDesc = rdr3["PIDesc"].ToString();
                    subjects.PIDet = rdr3["PIDet"].ToString();
                    subjects.PIType = rdr3["PIType"].ToString();
                    if (rdr3["PIDesc"].ToString().Equals("StopAt"))
                    {
                        subjects.PIDesc = "Stop Randomization At";
                    }
                    stopat.Add(subjects);
                }
                rdr3.Close();
                SqlCommand settingSelect = new SqlCommand(sql4, con);
                
                //settingSelect.Parameters.AddWithValue("@SPKEY", SPKEY);
                SqlDataReader rdr4 = settingSelect.ExecuteReader();
                while (rdr4.Read())
                {
                    var subjects = new SiteScreenStatus();
                    subjects.PIKEY = (int)rdr4["PIKEY"];
                    subjects.PIDesc = rdr4["PIDesc"].ToString();
                    subjects.PIDet = rdr4["PIDet"].ToString();
                    subjects.PIType = rdr4["PIType"].ToString();
                    
                  
                    sitescr.Add(subjects);
                    //ViewBag.randEmail = rdr4["PIDet"].ToString();
                }
                rdr4.Close();
                SqlCommand Shipemail = new SqlCommand(sql5, con);

                //Shipemail.Parameters.AddWithValue("@SPKEY", SPKEY);
                SqlDataReader rdr5 = Shipemail.ExecuteReader();
                while (rdr5.Read())
                {
                    var subjects = new SiteRandStatus();
                    subjects.PIKEY = (int)rdr5["PIKEY"];
                    subjects.PIDesc = rdr5["PIDesc"].ToString();
                    subjects.PIDet = rdr5["PIDet"].ToString();
                    subjects.PIType = rdr5["PIType"].ToString();

                    siternd.Add(subjects);
                    //ViewBag.shipEmail = rdr5["PIDet"].ToString();
                }
                rdr5.Close();
                SqlDataReader rdr6 = cmd6.ExecuteReader();
                while (rdr6.Read())
                {
                    var subjects = new StopScreenAt();
                    subjects.PIKEY = (int)rdr6["PIKEY"];
                    subjects.PIDesc = rdr6["PIDesc"].ToString();
                    subjects.PIDet = rdr6["PIDet"].ToString();
                    subjects.PIType = rdr6["PIType"].ToString();
                    if (rdr6["PIDesc"].ToString().Equals("StopScreenAt"))
                    {
                        subjects.PIDesc = "Stop Screening At";
                    }
                    stopscrn.Add(subjects);
                }
                rdr6.Close();

            }
            con.Close();
            list.randreq = randList;
            list.screenreq = screenstat;
            list.stopat = stopat;
            list.SiteScreen = sitescr;
            list.SiteRand = siternd;
            list.StopScreenAt = stopscrn;



            return View(list);
        }


        [HttpGet]
        public ActionResult GetUpdateClin(int? PIKEY)
        {
            Randstat model = new Randstat();

            if (PIKEY.HasValue)
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new SqlConnection(connectionString);
                string select = "SELECT PIKEY, PIDesc, PIDet, PIType FROM ProfileInfo  where PIKEY = @PIKEY";
                SqlCommand cmd = new SqlCommand(select, conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@PIKEY", PIKEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    
                    model.PIDesc = rdr["PIDesc"].ToString();
                    model.PIDet = rdr["PIDet"].ToString();
                    model.PIType = rdr["PIType"].ToString();
                }
                rdr.Close();
                conn.Close();
            }
            return PartialView("_EditRandClin", model);

        }

        public IActionResult UpdateRandStatus(int PIKEY, string PIDet)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            string sql = "Update ProfileInfo set PIDet = @PIDet,  CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER where PIKEY = @PIKEY";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            using (conn)
            {
                cmd.Parameters.AddWithValue("@PIDet", PIDet ); 
                cmd.Parameters.AddWithValue("@PIKEY", PIKEY);
                cmd.Parameters.AddWithValue("@CHANGEUSER", userid);
                cmd.ExecuteReader();
            }
            TempData["Message"] = "Information updated";
            return RedirectToAction("RequestHome");
        }

        [HttpGet]
        public ActionResult GetUpdateScreen(int? PIKEY)
        {
            ScreenInfo model = new ScreenInfo();

            if (PIKEY.HasValue)
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new SqlConnection(connectionString);
                string select = "SELECT PIKEY, PIDesc, PIDet, PIType FROM ProfileInfo  where PIKEY = @PIKEY";
                SqlCommand cmd = new SqlCommand(select, conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@PIKEY", PIKEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {

                    model.PIDesc = rdr["PIDesc"].ToString();
                    model.PIDet = rdr["PIDet"].ToString();
                    model.PIType = rdr["PIType"].ToString();
                }
                rdr.Close();
                conn.Close();
            }
            return PartialView("_EditScreenClin", model);

        }

        public ActionResult GetUpdateScrnStop(int? PIKEY)
        {
            ProfileInfo model = new ProfileInfo();

            if (PIKEY.HasValue)
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new SqlConnection(connectionString);
                string select = "SELECT PIKEY, PIDesc, PIDet, PIType FROM ProfileInfo  where PIKEY = @PIKEY";
                SqlCommand cmd = new SqlCommand(select, conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@PIKEY", PIKEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {

                    model.PIDesc = rdr["PIDesc"].ToString();
                    model.PIDet = rdr["PIDet"].ToString();
                    model.PIType = rdr["PIType"].ToString();
                }
                rdr.Close();
                conn.Close();
            }
            return PartialView("_EditStopAt", model);

        }


        public string GetEmail(string userID)
        {
            string sql = "SELECT zSecurityID.User_Email FROM zSecurityID WHERE (UserID = '" + userID + "') ";
            connectionString = _configuration.GetConnectionString("AmarexDbConnStr");
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    
                    object result = cmd.ExecuteScalar();
                    return result.ToString();
                }
            }

            
        }

        public string GetRequesterEmail(int prlbkey)
        {
            string sql = "SELECT ADDUSER from BIL_PRBLM WHERE  (PRBLMKEY = '" + prlbkey + "') ";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {

                    object result = cmd.ExecuteScalar();
                    return result.ToString();
                }
            }


        }
        
        ///Ship To Sites Sub Module
        public IActionResult ShiptoSites()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            //string userid = GetUserID();
            // string SITEID = GetSiteID(stringconnection, SPKEY, userid);
            string sql = "SELECT STSKEY, SITEID, INVNAME, SITENAME, ADDDATE, ADDUSER, ADDR1, ADDR2, CITY, STATE, ZIPCODE, COUNTRY FROM ShipToSite WHERE SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            //SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            var model = new List<SiteInfo>();

            using (conn)
            {
                conn.Open();
                //cmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var subjects = new SiteInfo();
                    subjects.STSKEY = (int)rdr["STSKEY"];
                    subjects.SITEID = rdr["SITEID"].ToString();
                    subjects.INVNAME = rdr["INVNAME"].ToString();
                    subjects.SITENAME = rdr["SITENAME"].ToString();
                    subjects.ADDDATE = (DateTime)rdr["ADDDATE"];
                    subjects.ADDUSER = rdr["ADDUSER"].ToString();
                    subjects.ADDR1 = rdr["ADDR1"].ToString();
                    subjects.ADDR2 = rdr["ADDR2"].ToString();
                    subjects.CITY = rdr["CITY"].ToString();
                    subjects.ZIPCODE = rdr["ZIPCODE"].ToString();
                    subjects.COUNTRY = rdr["COUNTRY"].ToString();
                    subjects.Location = rdr["ADDR1"].ToString() + " , " + rdr["CITY"].ToString();

                    model.Add(subjects);
                }

            }
            conn.Close();
            return View(model);
            
        }

        public SiteInfo GetSiteinfo(int STSKEY)
        {
            SiteInfo tempRep = new SiteInfo();
            tempRep.STSKEY = STSKEY;
            string sql = "SELECT STSKEY, SITEID, INVNAME, SITENAME, ADDR1, ADDR2, CITY, STATE, ZIPCODE, COUNTRY, PHONE, FAX, EMAIL, SpecialInstructions, AMAREX_COMM FROM ShipToSite WHERE STSKEY = @STSKEY";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@STSKEY", STSKEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    tempRep.STSKEY = (int)rdr["STSKEY"];
                    tempRep.SITEID = rdr["SITEID"].ToString();
                    tempRep.INVNAME = rdr["INVNAME"].ToString();
                    tempRep.SITENAME = rdr["SITENAME"].ToString();
                    tempRep.ADDR1 = rdr["ADDR1"].ToString();
                    tempRep.ADDR2 = rdr["ADDR2"].ToString();
                    tempRep.CITY = rdr["CITY"].ToString();
                    tempRep.COUNTRY = rdr["COUNTRY"].ToString();
                    if(tempRep.COUNTRY == "United States")
                    {
                        tempRep.USSTATE = rdr["STATE"].ToString();
                    }
                    else
                    {
                        tempRep.STATE = rdr["STATE"].ToString();
                    }
                    
                    tempRep.ZIPCODE = rdr["ZIPCODE"].ToString();
                    tempRep.PHONE = rdr["PHONE"].ToString();
                    tempRep.FAX = rdr["FAX"].ToString();
                    tempRep.EMAIL = rdr["EMAIL"].ToString();
                   
                    tempRep.SpecialInstructions = rdr["SpecialInstructions"].ToString();
                    tempRep.AMAREX_COMM = rdr["AMAREX_COMM"].ToString();
                }
                rdr.Close();


            }
            return tempRep;

        }

        [HttpGet]
        public IActionResult EditSite(int STSKEY)
        {
            // StudyDropdown();
            return View(GetSiteinfo(STSKEY));
        }

        [HttpPost]
        public IActionResult EditSite(int STSKEY, string SITENAME, string SITEID, string INVNAME, string ADDR1, string ADDR2, string CITY, string STATE, string USSTATE, string ZIPCODE, string PHONE, string FAX, string EMAIL, string SpecialInstructions, string AMAREX_COMM, string COUNTRY)
        {
           
           
            string userid = HttpContext.Session.GetString("suserid");
           
            string sql = "Update ShipToSite set SITEID = @SITEID, CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER, SITENAME = @SITENAME, INVNAME = @INVNAME, ADDR1 = @ADDR1, ADDR2 = @ADDR2, CITY = @CITY, STATE = @STATE, ZIPCODE = @ZIPCODE, COUNTRY = @COUNTRY, PHONE = @PHONE, FAX = @FAX, EMAIL = @EMAIL, SpecialInstructions = @SpecialInstructions, AMAREX_COMM = @AMAREX_COMM  where STSKEY = @STSKEY";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            //if (COUNTRY != "USA" && COUNTRY != "US" && COUNTRY != "United States of America" && COUNTRY != "U.S.")
            //{
            //    TempData["ErrorMessage"] = "Site address is not located in the U.S.";
            //    return View();
            //}
            if (COUNTRY == "United States" && USSTATE != null)
            {
                STATE = USSTATE;
            }
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();

                cmd.Parameters.AddWithValue("@STSKEY", STSKEY);
                cmd.Parameters.AddWithValue("@SITEID", SITEID);
                cmd.Parameters.AddWithValue("@SITENAME", SITENAME);
                cmd.Parameters.AddWithValue("@INVNAME", INVNAME);
                cmd.Parameters.AddWithValue("@ADDR1", ADDR1);
                cmd.Parameters.AddWithValue("@ADDR2", (object)ADDR2 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@CITY", CITY);
                cmd.Parameters.AddWithValue("@STATE", STATE);
                cmd.Parameters.AddWithValue("@ZIPCODE", ZIPCODE);
                cmd.Parameters.AddWithValue("@PHONE", PHONE);
                cmd.Parameters.AddWithValue("@FAX", (object)FAX ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@EMAIL", EMAIL);
                cmd.Parameters.AddWithValue("@COUNTRY", COUNTRY);
                cmd.Parameters.AddWithValue("@SpecialInstructions", (object)SpecialInstructions ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@AMAREX_COMM", (object)AMAREX_COMM ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@CHANGEUSER", userid);
                cmd.ExecuteNonQuery();
            }
            
            TempData["Message"] = "Site ID: " + SITEID + " information updated";
 
            return RedirectToAction("ShiptoSites");
        }

        public void updateSiteChangeUser(int STSKEY)
        {
            string sql = "Update ShipToSite Set CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER where STSKEY = @key";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            conn.Open();
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@key", STSKEY);
            cmd.Parameters.AddWithValue("@CHANGEUSER", HttpContext.Session.GetString("suserid"));
            cmd.ExecuteNonQuery();
        }

        public IActionResult DeleteSite(int STSKEY, string SITEID)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            updateSiteChangeUser(STSKEY);
            string sql = "Delete from ShipToSite where STSKEY = " + STSKEY;
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                cmd.ExecuteNonQuery();
            }
            TempData["Message"] = "Site ID: " + SITEID + " Deleted";
            return RedirectToAction("ShiptoSites");
        }

        

        public IActionResult AddSite()
        {
            return View();
        }
       [HttpPost]
        public IActionResult AddSite(SiteInfo Request, string USState)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string selectSql = "SELECT COUNT(*) FROM ShipToSite WHERE SITEID = @SITEID AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + "";
            string insertSql = "INSERT INTO ShipToSite (SPKEY, SITEID, INVNAME, SITENAME, ADDR1, ADDR2, CITY, STATE, ZIPCODE, COUNTRY, PHONE, FAX, EMAIL, SpecialInstructions, AMAREX_COMM, ADDUSER, ADDDATE) VALUES (@SPKEY, @SITEID, @INVNAME, @SITENAME, @ADDR1, @ADDR2, @CITY, @STATE, @ZIPCODE, @COUNTRY, @PHONE, @FAX, @EMAIL, @SpecialInstructions, @AMAREX_COMM, @ADDUSER, @ADDDATE)";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand selectCmd = new SqlCommand(selectSql, con);
            SqlCommand cmd = new SqlCommand(insertSql, con);
            using (con)
            {
                con.Open();
                
                selectCmd.Parameters.AddWithValue("@SITEID", Request.SITEID);
                //selectCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                int count = (int)selectCmd.ExecuteScalar();
                if (count > 0)
                {
                    // Handle invalid subjid
                    TempData["ErrorMessage"] = "Site has already been Registered";
                    return RedirectToAction("ShiptoSites");
                }

                //if(Request.COUNTRY != "USA" && Request.COUNTRY != "US" && Request.COUNTRY != "United States of America" && Request.COUNTRY != "U.S.")
                //{
                //    TempData["ErrorMessage"] = "Site address is not located in the U.S.";
                //    return View();
                //}
                if(Request.COUNTRY == "United States" && Request.USSTATE != null)
                {
                    Request.STATE = Request.USSTATE;
                }

                cmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                cmd.Parameters.AddWithValue("@SITEID", Request.SITEID);
                cmd.Parameters.AddWithValue("@SITENAME", Request.SITENAME);
                cmd.Parameters.AddWithValue("@INVNAME", Request.INVNAME);
                cmd.Parameters.AddWithValue("@ADDR1", Request.ADDR1);
                cmd.Parameters.AddWithValue("@ADDR2", (object)Request.ADDR2 ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@CITY", Request.CITY);
                cmd.Parameters.AddWithValue("@STATE", Request.STATE);
                cmd.Parameters.AddWithValue("@ZIPCODE", Request.ZIPCODE);
                cmd.Parameters.AddWithValue("@PHONE", Request.PHONE);
                cmd.Parameters.AddWithValue("@FAX", (object)Request.FAX ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@EMAIL", Request.EMAIL);
                cmd.Parameters.AddWithValue("@COUNTRY", Request.COUNTRY);
                cmd.Parameters.AddWithValue("@SpecialInstructions", (object)Request.SpecialInstructions ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@AMAREX_COMM", (object)Request.AMAREX_COMM ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@ADDUSER", userid);
                DateTime currentDateTime = DateTime.Now;
                cmd.Parameters.AddWithValue("@ADDDATE", currentDateTime);
                cmd.ExecuteNonQuery();
                string Message = "Site ID: "+Request.SITEID + " has been added.";
                TempData["Message"] = Message;
            }
            con.Close();
            

            return RedirectToAction("ShiptoSites");
        }


        public IActionResult IRTUsers()
        {

            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));

            //string userid = GetUserID();
            // string SITEID = GetSiteID(stringconnection, SPKEY, userid);
            connectionString = _configuration.GetConnectionString("AmarexDbConnStr");
            string sql = "SELECT zSecUserIDCenter.Center_Num as Site, zSecurityID.LName as lname, zSecurityID.FName as fname , zSecurityID.User_Email as email, zSecurityID.OrgName as orgname, zSecurityID.UserID as userid, zSecurityID.USER_TITLE as role FROM zSecurityID INNER JOIN zSecUserIDCenter ON zSecurityID.UserID = zSecUserIDCenter.UserID WHERE zSecUserIDCenter.SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY zSecUserIDCenter.Center_Num, zSecurityID.LName, zSecurityID.UserID";
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            var model = new List<IRTusers>();

            using (conn)
            {
                conn.Open();
                //cmd.Parameters.AddWithValue("@SPKEY", SPKEY);

                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var subjects = new IRTusers();
                    subjects.UserID = rdr["userid"].ToString();
                    subjects.LName = rdr["lname"].ToString();
                    subjects.FName = rdr["fname"].ToString();
                    subjects.OrgName = rdr["orgname"].ToString();
                    subjects.User_Email = rdr["email"].ToString();
                    subjects.Center_Num = rdr["Site"].ToString();
                    subjects.Role = rdr["role"].ToString();



                    model.Add(subjects);
                }

            }
            conn.Close();
            return View(model);

        }

        public IActionResult IRTUsersExcel(object sender, EventArgs e)
        {
            connectionString = _configuration.GetConnectionString("AmarexDbConnStr");
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            XLWorkbook wb = new XLWorkbook();
            DataTable dt = CreateDataTable("SELECT zSecUserIDCenter.Center_Num as SiteID, zSecurityID.LName as LastName, zSecurityID.FName as FirstName , zSecurityID.User_Email as Emailaddress, zSecurityID.USER_TITLE as Role, zSecurityID.OrgName as OrganizationName, zSecurityID.UserID as userid FROM zSecurityID INNER JOIN zSecUserIDCenter ON zSecurityID.UserID = zSecUserIDCenter.UserID WHERE (zSecUserIDCenter.SPKey = '" + SPKEY + "') ORDER BY zSecUserIDCenter.Center_Num, zSecurityID.LName, zSecurityID.UserID", connectionString);
            wb.Worksheets.Add(dt, "IRTUsersExcel").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "IRTUsersExcel" + DateTime.Now.Year + ".xlsx");
            }
        }

        public IActionResult SiteAddresses(object sender, EventArgs e)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            XLWorkbook wb = new XLWorkbook();
            DataTable dt = CreateDataTable("SELECT * FROM ShipToSite WHERE  SPKEY = " +SPKEY+ "", connectionString);
            wb.Worksheets.Add(dt, "SiteAddresses").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    " SiteAddresses" + DateTime.Now.Year + ".xlsx");
            }
        }

        public string GetUserEmail()
        {
            string userid = HttpContext.Session.GetString("suserid");
            string email = "";
            string sql = "SELECT User_Email FROM zSecurityID WHERE (UserID = '" + HttpContext.Session.GetString("suserid") + " ')";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("AmarexDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    email = rdr["User_Email"].ToString();
                }
            }
            return email;
        }


        public IActionResult EmailSettings()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            String randsql = "SELECT KEYID, PIDet, PIType, PIDESC, ADDDATE FROM Email_Notifications WHERE PIType = 'AddRandEmails' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            String shipsql = "SELECT KEYID, PIDet, PIType, PIDESC, ADDDATE  FROM Email_Notifications WHERE (PIType = 'ShipperEmail' OR PIType = 'PMIPLblRel'  OR PIType = 'PMIPLblRelUpdt' OR PIType = 'IPDepot') AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            String labelsql = "SELECT KEYID, PIDet, PIType, PIDESC, ADDDATE  FROM Email_Notifications WHERE PIType = 'Regulatory' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            String labelupsql = "SELECT KEYID, PIDet, PIType, PIDESC, ADDDATE  FROM Email_Notifications WHERE PIType = 'Emergency Unblinding' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            String biosql = "SELECT KEYID, PIDet, PIType, PIDESC, ADDDATE  FROM Email_Notifications WHERE PIType = 'PMIPLblKitListUL' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";

            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(randsql, con);
            SqlCommand cmd2 = new SqlCommand(shipsql, con);
            SqlCommand cmd3 = new SqlCommand(labelsql, con);
            SqlCommand cmd4 = new SqlCommand(labelupsql, con);
            SqlCommand cmd5 = new SqlCommand(biosql, con);


            var list = new Email_Notifications();
            var randList = new List<RandEmailNotif>();
            var screenstat = new List<ShipperEmail>();
            var Reg = new List<IPLabRelShip>();
            var Emer = new List<IPlabRelUp>();
            var bio = new List<BiometircUploadEmailNoti>();
         
            using (con)
            {
                con.Open(); 
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var subjects = new RandEmailNotif();
                    subjects.PIDet = rdr["PIDet"].ToString();
                    subjects.PIType = rdr["PIType"].ToString();
                    subjects.KEYID = (int)rdr["KEYID"];
                    subjects.ADDDATE = (DateTime)rdr["ADDDATE"];

                    randList.Add(subjects);
                }
                rdr.Close();
                SqlDataReader rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    var subjects = new ShipperEmail();
                    subjects.PIDet = rdr2["PIDet"].ToString();
                    subjects.PIType = rdr2["PIType"].ToString();
                    subjects.PIDESC = rdr2["PIDESC"].ToString();
                    subjects.KEYID = (int)rdr2["KEYID"];
                    subjects.ADDDATE = (DateTime)rdr2["ADDDATE"];

                    screenstat.Add(subjects);
                }
                rdr2.Close();
                SqlDataReader rdr3 = cmd3.ExecuteReader();
                while (rdr3.Read())
                {
                    var subjects = new IPLabRelShip();
                    subjects.PIDet = rdr3["PIDet"].ToString();
                    subjects.PIType = rdr3["PIType"].ToString();
                    subjects.KEYID = (int)rdr3["KEYID"];
                    subjects.ADDDATE = (DateTime)rdr3["ADDDATE"];

                    Reg.Add(subjects);
                }
                rdr3.Close();

                SqlDataReader rdr4 = cmd4.ExecuteReader();
                while (rdr4.Read())
                {
                    var subjects = new IPlabRelUp();
                    subjects.PIDet = rdr4["PIDet"].ToString();
                    subjects.PIType = rdr4["PIType"].ToString();
                    subjects.KEYID = (int)rdr4["KEYID"];
                    subjects.ADDDATE = (DateTime)rdr4["ADDDATE"];

                    Emer.Add(subjects);
                }
                rdr4.Close();

                SqlDataReader rdr5 = cmd5.ExecuteReader();
                while (rdr5.Read())
                {
                    var subjects = new BiometircUploadEmailNoti();
                    subjects.PIDet = rdr5["PIDet"].ToString();
                    subjects.PIType = rdr5["PIType"].ToString();
                    subjects.KEYID = (int)rdr5["KEYID"];
                    subjects.ADDDATE = (DateTime)rdr5["ADDDATE"];

                    bio.Add(subjects);
                }
                rdr5.Close();

            }
            con.Close();
            list.randemail = randList;
            list.shipemail = screenstat;
            list.Regulatory = Reg;
            list.EmergencyUnblind = Emer;
            list.Bioemail = bio;
            return View(list);
        }


        public IActionResult PostIP(string email, string type)
        {
            string[] arr = email.Split(",");
            for (int i = 0; i < arr.Count(); i++)
            {
               InsertNotif(arr[i].ToString(), type);
                string[] arrType = type.Split(".");
                type = arrType[0];
                string desc = arrType[1];
                string message = "This email is to inform you that you will be receiving " + desc.TrimEnd() + " emails. For any questions or concerns, please contact the Project Manager.";
                string subject = "Webview RTSM -  Notifcations";
                SendEmail(arr[i], subject, message);
            }
            TempData["Message"] = "Email Notifications assigned";
            return RedirectToAction("EmailSettings");
        }

        public void InsertNotif(string email, string type)
        {
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string sql = "insert into Email_Notifications ( SPKEY, PIType, PIDet, PIDESC, UserID, Enable, ADDDATE, ADDUSER) values (@SPKEY, @PIType, @PIDet, @PIDESC, @UserID, @Enable, @ADDDATE, @ADDUSER)";
            conn.Open();
            string[] arr = type.Split(".");
            type = arr[0];
            string desc = arr[1];
            string userId = GetUserID(email);
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@PIDet", email);
            cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
            cmd.Parameters.AddWithValue("@PIType", type);
            cmd.Parameters.AddWithValue("@PIDESC", desc);
            cmd.Parameters.AddWithValue("@UserID", (object)userId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Enable", 1);
            cmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
            cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));


            cmd.ExecuteNonQuery();

            conn.Close();
        }
        //public IActionResult PostShipperEmail(string email)
        //{
        //    string[] arr = email.Split(",");

        //    SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
        //    string sql = "insert into Email_Notifications ( SPKEY, PIType, PIDet, PIDESC, UserID, Enable) values (@SPKEY, @PIType, @PIDet, @PIDESC, @UserID, @Enable)";
        //    conn.Open();

        //    for (int i = 0; i < arr.Count(); i++)
        //    {
        //        string userId = GetUserID(arr[i]);
        //        SqlCommand cmd = new SqlCommand(sql, conn);
        //        cmd.Parameters.AddWithValue("@PIDet", arr[i]);
        //        cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
        //        cmd.Parameters.AddWithValue("@PIType", "ShipperEmail");
        //        cmd.Parameters.AddWithValue("@PIDESC", "Shipper");
        //        cmd.Parameters.AddWithValue("@UserID", (object)userId ?? DBNull.Value);
        //        cmd.Parameters.AddWithValue("@Enable", 1);

        //        cmd.ExecuteNonQuery();
        //    }

        //    conn.Close();
        //    return RedirectToAction("EmailSettings");
        //}

        public IActionResult PostRandEmail(string email)
        {
            string[] arr = email.Split(",");


            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string sql = "insert into Email_Notifications (SPKEY, PIType, PIDet, PIDESC, UserID, Enable, ADDDATE, ADDUSER) values (@SPKEY, @PIType, @PIDet, @PIDESC, @UserID, @Enable, @ADDDATE, @ADDUSER)";
            conn.Open();

            for (int i = 0; i < arr.Count(); i++)
            {
               
                    string userId = GetUserID(arr[i]);
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@PIDet", arr[i]);
                    cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
                    cmd.Parameters.AddWithValue("@PIType", "AddRandEmails");
                    cmd.Parameters.AddWithValue("@PIDESC", "Randomization");
                    cmd.Parameters.AddWithValue("@UserID", (object)userId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@Enable", 1);
                    cmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));
                cmd.ExecuteNonQuery();
                string message = "This email is to inform you that you will be receiving Randomization emails. For any questions or concerns, please contact the Project Manager.";
                string subject = "Webview RTSM -  Notifcations";
                SendEmail(arr[i], subject, message);

            }

            conn.Close();
            TempData["Message"] = "Email Notifications assigned";
            return RedirectToAction("EmailSettings");
        }

        //public IActionResult PostIpRelease(string email)
        //{
        //    string[] arr = email.Split(",");

        //    SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
        //    string sql = "insert into Email_Notifications (SPKEY, PIType, PIDet, PIDESC, UserID, Enable) values (@SPKEY, @PIType, @PIDet, @PIDESC, @UserID, @Enable)";
        //    conn.Open();

        //    for (int i = 0; i < arr.Count(); i++)
        //    {
        //        string userId = GetUserID(arr[i]);
        //        SqlCommand cmd = new SqlCommand(sql, conn);
        //        cmd.Parameters.AddWithValue("@PIDet", arr[i]);
        //        cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
        //        cmd.Parameters.AddWithValue("@PIType", "PMIPLblRel");
        //        cmd.Parameters.AddWithValue("@PIDESC", "IP LABELING RELEASE/SHIP");
        //        cmd.Parameters.AddWithValue("@UserID", (object)userId ?? DBNull.Value);
        //        cmd.Parameters.AddWithValue("@Enable", 1);
        //        cmd.ExecuteNonQuery();
        //    }

        //    conn.Close();
        //    return RedirectToAction("EmailSettings");
        //}

        public IActionResult PostIPBio(string email)
        {
            string[] arr = email.Split(",");

            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string sql = "insert into Email_Notifications (SPKEY, PIType, PIDet, PIDESC, UserID, Enable, ADDDATE, ADDUSER) values (@SPKEY, @PIType, @PIDet, @PIDESC, @UserID, @Enable, @ADDDATE, @ADDUSER)";
            conn.Open();

            for (int i = 0; i < arr.Count(); i++)
            {
                string userId = GetUserID(arr[i]);
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@PIDet", arr[i]);
                cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
                cmd.Parameters.AddWithValue("@PIType", "PMIPLblKitListUL");
                cmd.Parameters.AddWithValue("@PIDESC", "IP Biometrics Kit List Upload");
                cmd.Parameters.AddWithValue("@UserID", (object)userId ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Enable", 1);
                cmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));
                cmd.ExecuteNonQuery();

                string message = "This email is to inform you that you will be receiving IP Biometrics Kit List Upload emails. For any questions or concerns, please contact the Project Manager.";
                string subject = "Webview RTSM -  Notifcations";
                SendEmail(arr[i], subject, message);
            }

            conn.Close();
            TempData["Message"] = "Email Notifications assigned";
            return RedirectToAction("EmailSettings");
        }

        public IActionResult PostReg(string email)
        {
            string[] arr = email.Split(",");

            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string sql = "insert into Email_Notifications (SPKEY, PIType, PIDet, PIDESC, UserID, Enable, ADDDATE, ADDUSER) values (@SPKEY, @PIType, @PIDet, @PIDESC, @UserID, @Enable, @ADDDATE, @ADDUSER)";
            conn.Open();

            for (int i = 0; i < arr.Count(); i++)
            {
                string userId = GetUserID(arr[i]);
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@PIDet", arr[i]);
                cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
                cmd.Parameters.AddWithValue("@PIType", "Regulatory");
                cmd.Parameters.AddWithValue("@PIDESC", "Regulatory");
                cmd.Parameters.AddWithValue("@UserID", (object)userId ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Enable", 1);
                cmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));
                cmd.ExecuteNonQuery();

                string message = "This email is to inform you that you will be receiving Regulatory emails. For any questions or concerns, please contact the Project Manager.";
                string subject = "Webview RTSM -  Notifcations";
                SendEmail(arr[i], subject, message);
            }

            conn.Close();
            TempData["Message"] = "Email Notifications assigned";
            return RedirectToAction("EmailSettings");
        }

        public IActionResult PostEmerEmail(string email)
        {
            string[] arr = email.Split(",");

            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string sql = "insert into Email_Notifications (SPKEY, PIType, PIDet, PIDESC, UserID, Enable, ADDDATE, ADDUSER) values (@SPKEY, @PIType, @PIDet, @PIDESC, @UserID, @Enable, @ADDDATE, @ADDUSER)";
            conn.Open();

            for (int i = 0; i < arr.Count(); i++)
            {
                string userId = GetUserID(arr[i]);
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@PIDet", arr[i]);
                cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
                cmd.Parameters.AddWithValue("@PIType", "Emergency Unblinding");
                cmd.Parameters.AddWithValue("@PIDESC", "Emergency Unblinding");
                cmd.Parameters.AddWithValue("@UserID", (object)userId ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@Enable", 1);
                cmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));
                cmd.ExecuteNonQuery();

                string message = "This email is to inform you that you will be receiving Emergency Unblinding emails. For any questions or concerns, please contact the Project Manager.";
                string subject = "Webview RTSM -  Notifcations";
                SendEmail(arr[i], subject, message);
            }

            conn.Close();
            TempData["Message"] = "Email Notifications assigned";
            return RedirectToAction("EmailSettings");
        }

        //public IActionResult PostIPLabelUp(string email)
        //{
        //    string[] arr = email.Split(",");

        //    SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
        //    string sql = "insert into Email_Notifications (SPKEY, PIType, PIDet, PIDESC, UserID, Enable) values (@SPKEY, @PIType, @PIDet, @PIDESC, @UserID, @Enable)";
        //    conn.Open();


        //    for (int i = 0; i < arr.Count(); i++)
        //    {
        //        string userId = GetUserID(arr[i]);
        //        SqlCommand cmd = new SqlCommand(sql, conn);
        //        cmd.Parameters.AddWithValue("@PIDet", arr[i]);
        //        cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
        //        cmd.Parameters.AddWithValue("@PIType", "PMIPLblRelUpdt");
        //        cmd.Parameters.AddWithValue("@PIDESC", "IP LABELING RELEASE Updated");
        //       // cmd.Parameters.AddWithValue("@UserID", userId);
        //        cmd.Parameters.AddWithValue("@UserID", (object)userId ?? DBNull.Value);
        //        cmd.Parameters.AddWithValue("@Enable", 1);


        //        cmd.ExecuteNonQuery();
        //    }

        //    conn.Close();
        //    return RedirectToAction("EmailSettings");
        //}
        public IActionResult RemoveEmail(int KEYID)
        {
            string Notiftype = GetNotifType(KEYID);
            string[] arr = Notiftype.Split(";");
            updateChangeUser(KEYID);
            string sql = "delete from Email_Notifications where KEYID = @key";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            conn.Open();
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@key", KEYID);
            cmd.ExecuteNonQuery();
            conn.Close();
            string type = arr[0].TrimEnd();
            string email = arr[1].TrimEnd();
            //string[] arr = Notiftype.Split(";");
            string message = "This email is to inform you that you will no longer receive " + type + " emails. For any errors, please contact the Project Manager.";
            string subject = "Webview RTSM -  Notifcations";
            SendEmail(arr[1], subject, message);
            TempData["Message"] = "Email Notifications removed";
            return RedirectToAction("EmailSettings");
        }

        public void updateChangeUser(int KEYID)
        {
            string sql = "Update Email_Notifications Set CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER where KEYID = @key";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            conn.Open();
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@key", KEYID);
            cmd.Parameters.AddWithValue("@CHANGEUSER", HttpContext.Session.GetString("suserid"));
            cmd.ExecuteNonQuery();
        }

        public string GetUserID(string email)
        {
            //string sql = "SELECT zSecurityID.UserID FROM zSecurityID WHERE (User_Email = @Email)";
            // string sql = "SELECT zSecurityID.UserID FROM zSecurityID WHERE (User_Email = @Email)";
            //AND SSO_ID IS NOT Null
            string sql = "SELECT zSecurityID.UserID FROM zSecurityID WHERE (User_Email = @Email) AND SSO_ID IS NOT Null";


             connectionString = _configuration.GetConnectionString("AmarexDbConnStr");
            //connectionString = _configuration.GetConnectionString("AmarexDb");


            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.Parameters.AddWithValue("@Email", email); // Use parameterized query to prevent SQL injection

                    string result = cmd.ExecuteScalar() as string;

                    return result; // This will be null if no user is found or a string with the UserID
                }
            }
        }



        public IActionResult InventoryHome()
        {
           // List<SiteID> model = new List<SiteID>();
            var list = new InventoryModel();
            var site = new List<SiteID>();
            var stopsup = new List<StopAutoSupply>();
            var depinv = new List<DepotInv>();
            var study = new List<StopAutoSupply>();
            string sql = "SELECT DISTINCT SITEID FROM ShipToSite WHERE SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID";
            string sql2 = "SELECT PIDet, KEYID, PIDESC, PIType FROM Stop_Auto_Supply WHERE PIType = 'StopAutoResupply' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            string sql3 = "SELECT Top 100 KitNumber, SITEID, KITSTAT, KITKEY, KITSET FROM BIL_IP_RANGE WHERE (SITEID IS NULL) OR (KITSTAT = 'Shipped') ORDER BY KitNumber";
            string sql4 = "SELECT PIDet, KEYID, PIDESC, PIType FROM Stop_Auto_Supply WHERE PIType = 'AutoResupply' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";

            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            SqlCommand cmd2 = new SqlCommand(sql2, conn);
            SqlCommand cmd3 = new SqlCommand(sql3, conn);
            SqlCommand cmd4 = new SqlCommand(sql4, conn);
            using (conn)
            {
                conn.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var siteinfo = new SiteID();
                   // siteinfo.key = (int)rdr["STSKEY"];
                    siteinfo.SITEID = rdr["SITEID"].ToString();

                    site.Add(siteinfo);

                }
                rdr.Close();
                SqlDataReader rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    var subjects = new StopAutoSupply();
                    subjects.KEYID = (int)rdr2["KEYID"];
                    subjects.PIDESC = rdr2["PIDESC"].ToString();
                    subjects.PIDet = rdr2["PIDet"].ToString();
                    subjects.PIType = rdr2["PIType"].ToString();

                    stopsup.Add(subjects);
                }
                rdr2.Close();
                SqlDataReader rdr3 = cmd3.ExecuteReader();
                while (rdr3.Read())
                {
                    var subjects = new DepotInv();
                    
                    subjects.KitNumber = rdr3["KitNumber"].ToString();
                    subjects.SITEID = rdr3["SITEID"].ToString();
                    subjects.KITSTAT = rdr3["KITSTAT"].ToString();
                    subjects.KITKEY = (int)rdr3["KITKEY"];
                    subjects.KITSET = rdr3["KITSET"].ToString();
                    
                    


                    depinv.Add(subjects);
                }
                rdr3.Close();
                SqlDataReader rdr4 = cmd4.ExecuteReader();
                while (rdr4.Read())
                {
                    var subjects = new StopAutoSupply();
                    subjects.KEYID = (int)rdr4["KEYID"];
                    subjects.PIDESC = rdr4["PIDESC"].ToString();
                    subjects.PIDet = rdr4["PIDet"].ToString();
                    subjects.PIType = rdr4["PIType"].ToString();

                    study.Add(subjects);
                }
                rdr4.Close();
            }
            conn.Close();
            list.Site = site;
            list.Stopauto = stopsup;
            list.Depinv = depinv;
            list.StudyResupply = study;
            return View(list);
        }
        [HttpGet]   
        public IActionResult Inventory(string SITEID)
        {

            List<Inventory> Inv = new List<Inventory>();

            
            string sql = "SELECT SITEID, KitNumber, KITTYPE, SENDBY, RECVDBY, KITSTAT, KITSET, ASSIGNED, KITKEY FROM BIL_IP_RANGE WHERE((SITEID = @SITEID) AND((ASSIGNED IS NULL) OR (ASSIGNED = 'Unacceptable')) AND (KITSTAT = 'Acceptable' OR KITSTAT = 'Unacceptable' OR KITSTAT = 'Retire')) ORDER BY SITEID, KitNumber, KITKEY";
           
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@SITEID", SITEID);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                Inventory temp = new Inventory();

                temp.SITEID = rdr["SITEID"].ToString();
                temp.KitNumber = rdr["KitNumber"].ToString();
                temp.KITTYPE = rdr["KITTYPE"].ToString();
                temp.SENDBY = rdr["SENDBY"].ToString();
                temp.RECVDBY = rdr["RECVDBY"].ToString();
                temp.KITSTAT = rdr["KITSTAT"].ToString();
                temp.ASSIGNED = rdr["ASSIGNED"].ToString();
                temp.KITKEY = (int)rdr["KITKEY"];
                temp.KITSET = (short)rdr["KITSET"];
                Inv.Add(temp);
            }
            TempData["SiteID"] = SITEID;
            conn.Close();
            return View(Inv);
        }

        [HttpGet]
        public IActionResult InventoryShipment(string SITEID)
        {

            List<Inventory> Inv = new List<Inventory>();


            string sql = "SELECT SITEID, KITTYPE, SENDBY, KITSET FROM BIL_IP_RANGE WHERE((SITEID = @SITEID) AND((ASSIGNED IS NULL) OR (ASSIGNED = 'Unacceptable')) AND (KITSTAT = 'Acceptable' OR KITSTAT = 'Unacceptable' OR KITSTAT = 'Retire')) group by [SITEID],[KITSET],[KITTYPE], [SENDBY] order by KITSET";

            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@SITEID", SITEID);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                Inventory temp = new Inventory();

                temp.SITEID = rdr["SITEID"].ToString();
                temp.KITTYPE = rdr["KITTYPE"].ToString();
                temp.SENDBY = rdr["SENDBY"].ToString();
                temp.KITSET = (short)rdr["KITSET"];
                Inv.Add(temp);
            }
            TempData["SiteID"] = SITEID;
            conn.Close();
            return View(Inv);
        }

        
        public IActionResult ViewInventory(int KITSET, string SITEID)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            ViewBag.SiteID = SITEID;
            ViewBag.KITSET = KITSET;

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            string sql = "SELECT KITKEY, KitNumber, SITEID, KITSTAT FROM BIL_IP_RANGE WHERE (KITSET = @KITSET) AND (ASSIGNED IS NULL) AND (KITSTAT = 'Acceptable' OR KITSTAT = 'Unacceptable') ";
            SqlCommand cmd = new SqlCommand(sql, conn);
            //ShipIP temp = new ShipIP();
            var list = new IPReporting();
            var shipip = new List<ShipIP>();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // Retrieve existing data
                using (SqlCommand selectCmd = new SqlCommand(sql, con))
                {
                    selectCmd.Parameters.AddWithValue("@KITSET", KITSET);

                    SqlDataReader rdr = selectCmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var temp = new ShipIP();
                        temp.KITKEY = (int)rdr["KITKEY"];
                        temp.KitNumber = rdr["KitNumber"].ToString();
                        temp.SITEID = rdr["SITEID"].ToString();
                        temp.KITSTAT = rdr["KITSTAT"].ToString();

                        shipip.Add(temp);
                    }
                    rdr.Close();
                }
                con.Close();
            }

            list.ShipIP = shipip;
            return View(list);
        }

        [HttpPost]
        public IActionResult ViewInventory(int KITSET, string SITEID, string SelectedKitKeys)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            ViewBag.SiteID = SITEID;
            ViewBag.KITSET = KITSET;

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            string sql = "SELECT KITKEY, KitNumber, SITEID, KITSTAT FROM BIL_IP_RANGE WHERE (KITSET = @KITSET) AND (ASSIGNED IS NULL) AND (KITSTAT = 'Acceptable' OR KITSTAT = 'Unacceptable') ";
            SqlCommand cmd = new SqlCommand(sql, conn);
            //ShipIP temp = new ShipIP();
            var list = new IPReporting();
            var shipip = new List<ShipIP>();
            using (conn)
            {
                conn.Open();

                // Retrieve existing data
                using (SqlCommand selectCmd = new SqlCommand(sql, conn))
                {
                    selectCmd.Parameters.AddWithValue("@KITSET", KITSET);

                    SqlDataReader rdr = selectCmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var temp = new ShipIP();
                        temp.KITKEY = (int)rdr["KITKEY"];
                        temp.KitNumber = rdr["KitNumber"].ToString();
                        temp.SITEID = rdr["SITEID"].ToString();
                        temp.KITSTAT = rdr["KITSTAT"].ToString();

                        shipip.Add(temp);
                    }
                    rdr.Close();
                }
                conn.Close();
            }

            list.ShipIP = shipip;
            if (SelectedKitKeys == null)
            {
                TempData["ErrorMessage"] = "No Kit selected.";
                return View(list);
            }
            string[] kitKeysArray = SelectedKitKeys.Split(',');
            foreach (var kitKey in kitKeysArray)
            {
                string Upsql = "DECLARE @KC nvarchar(max), @KITS nvarchar(50); SELECT @KC = KITCOMM, @KITS = KITSTAT FROM BIL_IP_RANGE WHERE (KITKEY = " + kitKey + "); IF @KC IS NULL BEGIN SELECT @KC = 'Reset to SHIP from ' + @KITS + ' on ' + CAST(sysdatetime() AS nvarchar(50)) END ELSE BEGIN SELECT @KC = @KC + ' ---- Reset to SHIP from ' + @KITS + ' on ' + CAST(sysdatetime() AS nvarchar(50)) END; ";
                Upsql += "UPDATE BIL_IP_RANGE SET RECVDBY = null, RECVDDATE = null, KITSTAT = 'Shipped', ASSIGNED = null, isExcrusion = null, SPAPDTC = null, IsPhysicalDamage = null,  KITCOMM = @KC WHERE (KITKEY = " + kitKey + ")";

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                   
                    using (SqlCommand sqlcmd = new SqlCommand(Upsql, con))
                    {
                        con.Open();
                        sqlcmd.ExecuteNonQuery();
                    }
                    
                }
            }

            TempData["Message"] = "Selected kits have been reset to ship status for Site ID: " + SITEID;
            return RedirectToAction("InventoryShipment", new { SITEID = SITEID });
        }

        //Kit Level resetting status
        public IActionResult UpdateKitStatus(int kitKey, string SITEID, string KitNumber)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql = "DECLARE @KC nvarchar(max), @KITS nvarchar(50); SELECT @KC = KITCOMM, @KITS = KITSTAT FROM BIL_IP_RANGE WHERE (KITKEY = " + kitKey + "); IF @KC IS NULL BEGIN SELECT @KC = 'Reset to SHIP from ' + @KITS + ' on ' + CAST(sysdatetime() AS nvarchar(50)) END ELSE BEGIN SELECT @KC = @KC + ' ---- Reset to SHIP from ' + @KITS + ' on ' + CAST(sysdatetime() AS nvarchar(50)) END; ";
            sql += "UPDATE BIL_IP_RANGE SET RECVDBY = null, RECVDDATE = null, KITSTAT = 'Shipped', ASSIGNED = null, isExcrusion = null, SPAPDTC = null, IsPhysicalDamage = null,  KITCOMM = @KC WHERE (KITKEY = " + kitKey + ")";
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                cmd.ExecuteNonQuery();
            }
           
            TempData["Message"] = "Status of Kit Number: " + KitNumber + " updated for Site ID: " + SITEID;
           
            return RedirectToAction("Inventory", new { SITEID = SITEID });
        }

        public ActionResult GetAutoResupply(int? KEYID)
        {
            StopAutoSupply model = new StopAutoSupply();

            if (KEYID.HasValue)
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new SqlConnection(connectionString);
                string select = "SELECT KEYID, PIDESC, PIDet, PIType FROM Stop_Auto_Supply where KEYID = @KEYID";
                SqlCommand cmd = new SqlCommand(select, conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@KEYID", KEYID);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    model.PIDESC = rdr["PIDESC"].ToString();
                    model.PIDet = rdr["PIDet"].ToString();
                    model.PIType = rdr["PIType"].ToString();
                }
                rdr.Close();
                conn.Close();
            }
            return PartialView("_EditAutoSupply", model);

        }

        public IActionResult UpdateAutoResupply(int KEYID, string PIDet)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            string sql = "Update Stop_Auto_Supply set PIDet = @PIDet,  CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER where KEYID = @KEYID";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            using (conn)
            {
                cmd.Parameters.AddWithValue("@PIDet", PIDet);
                cmd.Parameters.AddWithValue("@KEYID", KEYID);
                cmd.Parameters.AddWithValue("@CHANGEUSER", userid);
                cmd.ExecuteReader();
            }
            TempData["Message"] = "Information updated";
            return RedirectToAction("InventoryHome");
        }
        public IActionResult PostSite(string SiteID)
        {
            string userid = HttpContext.Session.GetString("suserid");
            string[] arr = SiteID.Split(",");

            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string sql = "insert into Stop_Auto_Supply (SPKEY, PIType, PIDet, PIDESC, ADDDATE, ADDUSER) values (@SPKEY, @PIType, @PIDet, @PIDESC, @ADDDATE, @ADDUSER)";
            conn.Open();

            for (int i = 0; i < arr.Count(); i++)
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@PIDet", arr[i]);
                cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
                cmd.Parameters.AddWithValue("@PIType", "StopAutoResupply");
                cmd.Parameters.AddWithValue("@PIDESC", "StopAutoResupply");
                cmd.Parameters.AddWithValue("@ADDUSER", userid);
                cmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);

                cmd.ExecuteNonQuery();
            }

            conn.Close();
            TempData["Message"] = "Automatic Re-supply stopped for Site :" + SiteID + "";
            return RedirectToAction("InventoryHome");
           
        }


        public IActionResult RemoveSite(int KEYID, string SITEID)
        {
            string sql = "delete from Stop_Auto_Supply where KEYID = @key";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            conn.Open();
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@key", KEYID);
            cmd.ExecuteNonQuery();
            conn.Close();
            TempData["Message"] = "SiteID: " + SITEID + " Removed";

            return RedirectToAction("InventoryHome");
        }
        private IList<SiteInfo> LoadSites()
        {
     
            string sql = "SELECT DISTINCT SITEID FROM ShipToSite WHERE SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);

            IList<SiteInfo> events = new List<SiteInfo>();

            using (conn)
            {
                conn.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var site = new SiteInfo();

                    site.description = rdr["SITEID"].ToString();
                   
                    events.Add(site);

                }
            }
            conn.Close();

            return events;

        }

        public JsonResult GetAllSites()
        {
            try
            {
                IList<SiteInfo> eventslist = LoadSites();

                var modifiedData = eventslist.Select(d =>
             new
             {


                d.description,
               
             }



                 );

                return Json(modifiedData);
            }


            catch (Exception ex)
            {
                return Json(new { IsSuccess = false, Message = ex.Message });
            }


        }

        public IActionResult Resetshipdepot(int kitKey, string SITEID, string KitNumber, string kitstat)
        {
            if(kitstat == "Shipped")
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                string sql = "UPDATE BIL_IP_RANGE SET [SITEID] = null, [KITSET] = null, [KITTYPE] = null, [SENDBY] = null, [SENDDATE] = null, [KITSTAT] = null WHERE (KITKEY = " + kitKey + ")";
                SqlConnection conn = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand(sql, conn);
                using (conn)
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                string logdesc = "Reset to Depot from Ship";
                InsDataChng(HttpContext.Session.GetString("sesSPKey"), HttpContext.Session.GetString("suserid"),logdesc, "Reset to Depot from Ship" );
                TempData["Message"] = "Kit: " + KitNumber + " Reset to Depot from Ship";
                return RedirectToAction("InventoryHome");
            }
            else
            {
                TempData["ErrorMessage"] = "Can not process, IP/Kit is not in Ship status.";
                return RedirectToAction("InventoryHome");
            }
        }

        public IActionResult Damagedepot(int kitKey, string KitNumber, string SITEID)
        {

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql = "UPDATE BIL_IP_RANGE SET [SITEID] = 'Unacceptable', [KITSTAT] = 'Unacceptable', [KITTYPE] = 'Unacceptable',  KITCOMM = 'Damaged at depot ' + CAST(sysdatetime() AS nvarchar(50)) WHERE (KITKEY = " + kitKey + ")";
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                cmd.ExecuteNonQuery();
            }
            string logdesc = "Marked As Unacceptable";
           InsDataChng(HttpContext.Session.GetString("sesSPKey"), HttpContext.Session.GetString("suserid"), logdesc, "Marked Damage");
            TempData["ErrorMessage"] = "Kit: " + KitNumber + " Marked As Unacceptable";
            return RedirectToAction("InventoryHome");
        }


        public IActionResult ResetshipdepotView(int kitKey, string SITEID, string KitNumber, string kitstat)
        {
            if (kitstat == "Shipped")
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                string sql = "UPDATE BIL_IP_RANGE SET [SITEID] = null, [KITSET] = null, [KITTYPE] = null, [SENDBY] = null, [SENDDATE] = null, [KITSTAT] = null WHERE (KITKEY = " + kitKey + ")";
                SqlConnection conn = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand(sql, conn);
                using (conn)
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                string logdesc = "Reset to Depot from Ship";
                InsDataChng(HttpContext.Session.GetString("sesSPKey"), HttpContext.Session.GetString("suserid"), logdesc, "Reset to Depot from Ship");
                TempData["Message"] = "Kit: " + KitNumber + " Reset to Depot from Ship";
                return RedirectToAction("ViewIPDepot");
            }
            else
            {
                TempData["ErrorMessage"] = "Can not process, IP/Kit is not in Ship status.";
                return RedirectToAction("ViewIPDepot");
            }
        }

        public IActionResult DamagedepotView(int kitKey, string KitNumber, string SITEID)
        {

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql = "UPDATE BIL_IP_RANGE SET [SITEID] = 'Unacceptable', [KITSTAT] = 'Unacceptable', [KITTYPE] = 'Unacceptable', KITCOMM = 'Damaged at depot ' + CAST(sysdatetime() AS nvarchar(50)) WHERE (KITKEY = " + kitKey + ")";
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                cmd.ExecuteNonQuery();
            }
            string logdesc = "Marked As Unacceptable";
            InsDataChng(HttpContext.Session.GetString("sesSPKey"), HttpContext.Session.GetString("suserid"), logdesc, "Marked Damage");
            TempData["ErrorMessage"] = "Kit: " + KitNumber + " Marked As Unacceptable";
            return RedirectToAction("ViewIPDepot");
        }

        public void InsDataChng(string spkey, string logUid, string logDesc, string logType)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);

            SqlCommand logCmd = new SqlCommand("LogAppActIns", conn);
           logCmd.CommandType = CommandType.StoredProcedure;

            //    cmd.Parameters.Add("@pLogin", SqlDbType.VarChar).Value = User.FindFirst(ClaimTypes.NameIdentifier).Value;
            logCmd.Parameters.Add("@spkey", SqlDbType.Int).Value = spkey;
            logCmd.Parameters.Add("@loguser", SqlDbType.VarChar).Value = logUid;
            logCmd.Parameters.Add("@logdesc", SqlDbType.VarChar).Value = logDesc;
            logCmd.Parameters.Add("@typelog", SqlDbType.VarChar).Value = logType;


            using (conn)
            {
                conn.Open();
                logCmd.ExecuteReader();
            }
        }

        
            public static bool IsValidEmailPattern(string email)
            {
                try
                {
                    var mailAddress = new MailAddress(email);
                    return true;
                }
                catch (FormatException)
                {
                    return false;
                }
            }


        private IList<IRTusers> LoadEmails()
        {
            
            //string sql = "Select DISTINCT User_Email From zSecurityID WHERE User_Email is not null order by User_Email";
            string sql = "SELECT DISTINCT zSecurityID.User_Email As User_Email FROM zSecurityID INNER JOIN zSecUserIDCenter ON zSecurityID.UserID = zSecUserIDCenter.UserID WHERE zSecUserIDCenter.SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";

            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("AmarexDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);

            IList<IRTusers> events = new List<IRTusers>();

            using (conn)
            {
                conn.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var user = new IRTusers();

                    //user.Title = rdr["User_Email"].ToString();

                    //user.description = rdr["User_Email"].ToString();

                    user.Title = rdr["User_Email"].ToString();

                    user.description = rdr["User_Email"].ToString();

                    events.Add(user);

                }
            }
            conn.Close();

            return events;

        }

        public JsonResult GetAllEmails()
        {
            try
            {
                IList<IRTusers> eventslist = LoadEmails();

                var modifiedData = eventslist.Select(d =>
             new
             {


                 d.Title,
                 d.description,
                 
             }



                 );

                return Json(modifiedData);
            }


            catch (Exception ex)
            {
                return Json(new { IsSuccess = false, Message = ex.Message });
            }


        }

        public IActionResult EditEmail(int KEYID, string Email)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string userid = HttpContext.Session.GetString("suserid");
            string sql = "UPDATE Email_Notifications SET  PIDet = @PIDet, CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER WHERE (KEYID = " + KEYID + ")";
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@PIDet", Email);
                cmd.Parameters.AddWithValue("@CHANGEUSER", userid);
                cmd.ExecuteNonQuery();
            }
            TempData["Message"] = "Email Address Updated";
            return RedirectToAction("EmailSettings");       
        }


        [HttpGet]
        public ActionResult GetUpdateEmail(int? KEYID)
        {
            Email_Notifications model = new Email_Notifications();

            if (KEYID.HasValue)
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new SqlConnection(connectionString);
                string select = "SELECT PIDet FROM Email_Notifications  where KEYID = @KEYID";
                SqlCommand cmd = new SqlCommand(select, conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@KEYID", KEYID);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    model.PIDet = rdr["PIDet"].ToString();
                }
                rdr.Close();
                conn.Close();
            }
            return PartialView("_EditEmail", model);

        }

        public string GetNotifType(int KEYID)
        {
            string rtnVal ="";
            string sql = "SELECT PIDESC, PIDet FROM Email_Notifications WHERE (KEYID = " + KEYID + ")";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
               // cmd.Parameters.AddWithValue("@KEYID", KEYID);
                while (rdr.Read())
                {
                    rtnVal = rdr["PIDESC"].ToString() + ";" + rdr["PIDet"].ToString();
                }
            }
            return rtnVal;
        }


        public string checkidpwd(string username, string password)
        {

            var rtnVal = "";
            SecSSO chkSSO2 = new SecSSO();
            rtnVal = chkSSO2.ChkIDPWSSO(username, password, HttpContext.Session.GetString("sesuriSSIS"), HttpContext.Session.GetString("sesinstanceID"), HttpContext.Session.GetString("sesSecurityKey"), HttpContext.Session.GetString("sesAmarexDb"));
            return rtnVal;
        }

        private bool IsValidUser(string username, string password)
        {
            string check = checkidpwd(username, password);

            // Check if the provided username and password match the valid credentials
            if (check.Equals("7103"))
            {

                return true; // Valid user
            }
            else
            {
                return false; // Invalid user
            }
        }


        public IActionResult SiteScreen(string SiteID)
        {
            string[] arr = SiteID.Split(",");
            string userid = HttpContext.Session.GetString("suserid");
            //Check if Site is already added
            string rtnVal = "";
            rtnVal = CheckSite(SiteID, "SiteScreen");
            if (rtnVal == "Site Found")
            {
                TempData["ErrorMessage"] = "One or more Site(s) already added.";
                return RedirectToAction("RequestHome");
            }

            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string sql = "insert into  ProfileInfo (SPKEY, PIType, PIDet, PIDesc, ADDDATE, ADDUSER) values (@SPKEY, @PIType, @PIDet, @PIDesc, @ADDDATE, @ADDUSER)";
            conn.Open();

            for (int i = 0; i < arr.Count(); i++)
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@PIDet", arr[i]);
                cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
                cmd.Parameters.AddWithValue("@PIType", "SiteScreen");
                cmd.Parameters.AddWithValue("@PIDesc", "Screen Enabled");
                cmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));

                cmd.ExecuteNonQuery();
            }

            conn.Close();
            TempData["Message"] = "Screening enabled for Site :" + SiteID + "";
            return RedirectToAction("RequestHome");

        }

        public string CheckSite(string SiteID, string PIType)
        {
            string rtnVal = "";
            string[] arr = SiteID.Split(",");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            string sql = "SELECT COUNT(*) FROM ProfileInfo WHERE PIDET = @PIDET AND PIType = @PIType";
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                for (int i = 0; i < arr.Count(); i++)
                {
                    cmd.Parameters.AddWithValue("@PIDET", SiteID);
                    cmd.Parameters.AddWithValue("@PIType", PIType);

                    //selectCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                    int count = (int)cmd.ExecuteScalar();
                    if (count >= 1)
                    {
                        // Handle invalid subjid
                        rtnVal = "Site Found";
                        return rtnVal;

                    }
                }
                conn.Close();
            }
            return rtnVal;

        }

        public IActionResult SiteRand(string SiteID)
        {
            string[] arr = SiteID.Split(",");
            string rtnVal = "";
            rtnVal = CheckSite(SiteID, "SiteRand");
            ViewBag.Sites = SiteID;
            if (rtnVal == "Site Found")
            {
                TempData["ErrorMessage"] = "One or more Site(s) already added.";
                return RedirectToAction("RequestHome");
            }

            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string sql = "insert into  ProfileInfo (SPKEY, PIType, PIDet, PIDesc, ADDDATE, ADDUSER) values (@SPKEY, @PIType, @PIDet, @PIDesc, @ADDDATE, @ADDUSER)";
            conn.Open();

            for (int i = 0; i < arr.Count(); i++)
            {
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@PIDet", arr[i]);
                cmd.Parameters.AddWithValue("@SPKEY", HttpContext.Session.GetString("sesSPKey"));
                cmd.Parameters.AddWithValue("@PIType", "SiteRand");
                cmd.Parameters.AddWithValue("@PIDesc", "Rand Enabled");
                cmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));

                cmd.ExecuteNonQuery();
            }

            conn.Close();
            TempData["Message"] = "Randomization enabled for Site :" + SiteID + "";
            return RedirectToAction("RequestHome");

        }

        [HttpGet]
        public ActionResult GetUpdateSiteRand(int? PIKEY)
        {
            SiteRandStatus model = new SiteRandStatus();

            if (PIKEY.HasValue)
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new SqlConnection(connectionString);
                string select = "SELECT PIKEY, PIDesc, PIDet, PIType FROM ProfileInfo  where PIKEY = @PIKEY";
                SqlCommand cmd = new SqlCommand(select, conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@PIKEY", PIKEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    model.PIDesc = rdr["PIDesc"].ToString();
                    model.PIDet = rdr["PIDet"].ToString();
                    model.PIType = rdr["PIType"].ToString();
                }
                rdr.Close();
                conn.Close();
            }
            return PartialView("_EditSiteRand", model);

        }

        public IActionResult UpdateSiteRand(int PIKEY, string PIDesc)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            string sql = "Update ProfileInfo set PIDesc = @PIDesc,  CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER where PIKEY = @PIKEY";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            using (conn)
            {
                cmd.Parameters.AddWithValue("@PIDesc", PIDesc);
                cmd.Parameters.AddWithValue("@PIKEY", PIKEY);
                cmd.Parameters.AddWithValue("@CHANGEUSER", userid);
                cmd.ExecuteReader();
            }
            TempData["Message"] = "Information updated";
            return RedirectToAction("RequestHome");
        }

        [HttpGet]
        public ActionResult GetUpdateSiteScreen(int? PIKEY)
        {
            SiteScreenStatus model = new SiteScreenStatus();

            if (PIKEY.HasValue)
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new SqlConnection(connectionString);
                string select = "SELECT PIKEY, PIDesc, PIDet, PIType FROM ProfileInfo  where PIKEY = @PIKEY";
                SqlCommand cmd = new SqlCommand(select, conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@PIKEY", PIKEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    model.PIDesc = rdr["PIDesc"].ToString();
                    model.PIDet = rdr["PIDet"].ToString();
                    model.PIType = rdr["PIType"].ToString();
                }
                rdr.Close();
                conn.Close();
            }
            return PartialView("_EditSiteScreen", model);

        }

        public IActionResult UpdateSiteScreen(int PIKEY, string PIDesc)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            string sql = "Update ProfileInfo set PIDesc = @PIDesc,  CHANGEDATE = SYSDATETIME(), CHANGEUSER = @CHANGEUSER where PIKEY = @PIKEY";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            using (conn)
            {
                cmd.Parameters.AddWithValue("@PIDesc", PIDesc);
                cmd.Parameters.AddWithValue("@PIKEY", PIKEY);
                cmd.Parameters.AddWithValue("@CHANGEUSER", userid);
                cmd.ExecuteReader();
            }
            TempData["Message"] = "Information updated";
            return RedirectToAction("RequestHome");
        }


        public ActionResult GetUpdateStopScrn(int? PIKEY)
        {
            StopScreenAt model = new StopScreenAt();

            if (PIKEY.HasValue)
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new SqlConnection(connectionString);
                string select = "SELECT PIKEY, PIDesc, PIDet, PIType FROM ProfileInfo  where PIKEY = @PIKEY";
                SqlCommand cmd = new SqlCommand(select, conn);
                conn.Open();
                cmd.Parameters.AddWithValue("@PIKEY", PIKEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {

                    model.PIDesc = rdr["PIDesc"].ToString();
                    model.PIDet = rdr["PIDet"].ToString();
                    model.PIType = rdr["PIType"].ToString();
                }
                rdr.Close();
                conn.Close();
            }
            return PartialView("_EditStopScreenAt", model);

        }


        public IActionResult ReportsHome()
        {
            return View();
        }

        public IActionResult SubjectReport()
        { //since the study table does not currently contain all of the codes to allow for joining I full joined all the log tables on project codes and will just take the relavent info from each
          // string sql = "select * from PROJECT_TEAM_LOG a full join ClinProjectLogs b on a.PROJECTCODE = b.ProjectCode full join DMProjectLogs c on a.PROJECTCODE = c.AmaProjCode full join ITProjectLogs d on a.PROJECTCODE = d.ProjectCode full join MWProjectLogs e on a.PROJECTCODE = e.ProjectCode full join RAProjectLogs f on a.PROJECTCODE = f.ProjectCode full join SafetyProjectLogs g on a.PROJECTCODE = g.StudyCode";
            string sql = "SELECT SITEID, SUBJID, BRTHDTC, SEX, AgeGroup, ICDTC, SCRNDTC, STATUS_INFO, PMCOMDATE, ADDDATE, ADDUSER, DATE_RAND, RANDBY, SFDATE FROM BIL_SUBJ WHERE SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID   "; 
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Subject_Status" + DateTime.Now.ToString("MM-dd-yyyy"));


            worksheet.Cell(1, 1).Value = "Site ID";
            worksheet.Cell(1, 2).Value = "Subject ID";
            worksheet.Cell(1, 3).Value = "Year of Birth";
            worksheet.Cell(1, 4).Value = "Sex";
            worksheet.Cell(1, 5).Value = "Age Group";
            worksheet.Cell(1, 6).Value = "Informed Consent Date";
            worksheet.Cell(1, 7).Value = "Status";
            worksheet.Cell(1, 8).Value = "Screened Date";
            worksheet.Cell(1, 9).Value = "Screened By";
            worksheet.Cell(1, 10).Value = "Randomization Date";
            worksheet.Cell(1, 11).Value = "Randomized By";
           
            worksheet.Cell(1, 12).Value = "Screen Failed Date";
            //worksheet.Cell(1, 13).Value = "Lead Statistician";
            //worksheet.Cell(1, 14).Value = "Medical Monitor";
            //worksheet.Cell(1, 15).Value = "PV Monitor";
            //worksheet.Cell(1, 16).Value = "Authorized Representative";


            int i = 2;

            while (rdr.Read())
            {
                worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["SITEID"]) ? rdr["SITEID"].ToString() : null;
                worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["SUBJID"]) ? rdr["SUBJID"].ToString() : null;
                worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["BRTHDTC"]) ? rdr["BRTHDTC"].ToString() : null;
                worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["SEX"]) ? rdr["SEX"].ToString() : null;
                worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["AgeGroup"]) ? rdr["AgeGroup"].ToString() : null;
                worksheet.Cell(i, 6).Value = !Convert.IsDBNull(rdr["ICDTC"]) ? rdr["ICDTC"].ToString() : null;
                worksheet.Cell(i, 7).Value = !Convert.IsDBNull(rdr["STATUS_INFO"]) ? rdr["STATUS_INFO"].ToString() : null;
                worksheet.Cell(i, 8).Value = !Convert.IsDBNull(rdr["ADDDATE"]) ? rdr["ADDDATE"].ToString() : null;
                worksheet.Cell(i, 9).Value = !Convert.IsDBNull(rdr["ADDUSER"]) ? rdr["ADDUSER"].ToString() : null;
                worksheet.Cell(i, 10).Value = !Convert.IsDBNull(rdr["DATE_RAND"]) ? rdr["DATE_RAND"].ToString() : null;
                worksheet.Cell(i, 11).Value = !Convert.IsDBNull(rdr["RANDBY"]) ? rdr["RANDBY"].ToString() : null;
                worksheet.Cell(i, 12).Value = !Convert.IsDBNull(rdr["SFDATE"]) ? rdr["SFDATE"].ToString() : null;
                i++;

            }
            conn.Close();
            worksheet.Columns().AdjustToContents();
            return View(ImportExceltoDatatable(workbook));

        }

        public static DataTable ImportExceltoDatatable(XLWorkbook file)
        {
            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (file)
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = file.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;

                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }

                return dt;
            }



        }



        public static DataTable ImportExceltoDatatable2(XLWorkbook file)
        {
            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (file)
            {
                // Read the first Sheet from Excel file.
                IXLWorksheet workSheet = file.Worksheet(1);

                // Create a new DataTable.
                DataTable dt = new DataTable();

                // Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    // Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        // Add rows to DataTable.
                        dt.Rows.Add();

                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            // Ensure the cell is not null before reading its value
                            if (row.Cell(i + 1) != null)
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = row.Cell(i + 1).Value.ToString();
                            }
                            else
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = null;
                            }
                        }
                    }
                }

                return dt;
            }
        }



        public IActionResult SubjectReportExcel()
        {

            XLWorkbook wb = new XLWorkbook();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DataTable dt = CreateDataTable("SELECT SITEID AS 'Site ID', SUBJID AS 'Subject ID', BRTHDTC AS 'Year of Birth', SEX AS 'Sex', AgeGroup as 'Age Group', ICDTC AS 'Informed Consent Date', SCRNDTC As 'Screened Date', STATUS_INFO AS 'Status',  ADDDATE As 'Screened Date', ADDUSER AS 'Screened By', DATE_RAND AS 'Randomization Date', RANDBY AS 'Randomized By', SFDATE As 'Screen Failed Date ' FROM BIL_SUBJ WHERE SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID   ", connectionString);
            wb.Worksheets.Add(dt, "Subject Status Report").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Subject Status Report" + DateTime.Now.Year + ".xlsx");
            }

        }


        //public IActionResult SubjectVisitReport()
        //{
        //    string sql = "SELECT [VISIT], KitNumber, REPLACE(UPPER(CONVERT(varchar, [ASSIGNMENT_DATE], 106)), ' ', '/') AS 'Dispense', ASSIGNED_BY, ExpiryDate, LotNo, KitRepled, SITEID, SUBJID, notdone, ReaYes FROM BIL_IP_RANGE WHERE SITEID is not NULL AND SUBJID is not null ORDER BY SITEID, SUBJID";
        //    connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
        //    SqlConnection conn = new SqlConnection(connectionString);
        //    SqlCommand cmd = new SqlCommand(sql, conn);
        //    conn.Open();
        //    SqlDataReader rdr = cmd.ExecuteReader();

        //    var workbook = new XLWorkbook();
        //    var worksheet = workbook.Worksheets.Add("Subject_Visit" + DateTime.Now.ToString("MM-dd-yyyy"));


        //    worksheet.Cell(1, 1).Value = "Site ID";
        //    worksheet.Cell(1, 2).Value = "Subject ID";
        //    worksheet.Cell(1, 3).Value = "Visit";
        //    worksheet.Cell(1, 4).Value = "Kit Number";
        //    worksheet.Cell(1, 5).Value = "Dispensed Date";
        //    worksheet.Cell(1, 6).Value = "Dispensed By";
        //    worksheet.Cell(1, 7).Value = "Expiry Date";
        //    worksheet.Cell(1, 8).Value = "Lot Number";
        //    worksheet.Cell(1, 9).Value = "Replacement For";

        //    worksheet.Cell(1, 10).Value = "Visit Not done";
        //    worksheet.Cell(1, 11).Value = "Reason(IP dispensed, visit not done)";
        //    //worksheet.Cell(1, 15).Value = "PV Monitor";
        //    //worksheet.Cell(1, 16).Value = "Authorized Representative";


        //    int i = 2;

        //    while (rdr.Read())
        //    {
        //        worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["SITEID"]) ? rdr["SITEID"].ToString() : null;
        //        worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["SUBJID"]) ? rdr["SUBJID"].ToString() : null;
        //        worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["VISIT"]) ? rdr["VISIT"].ToString() : null;
        //        worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["KitNumber"]) ? rdr["KitNumber"].ToString() : null;
        //        worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["Dispense"]) ? rdr["Dispense"].ToString() : null;
        //        worksheet.Cell(i, 6).Value = !Convert.IsDBNull(rdr["ASSIGNED_BY"]) ? rdr["ASSIGNED_BY"].ToString() : null;
        //        worksheet.Cell(i, 7).Value = !Convert.IsDBNull(rdr["IPLblShipExpiryDate"]) ? rdr["IPLblShipExpiryDate"].ToString() : null;
        //        worksheet.Cell(i, 8).Value = !Convert.IsDBNull(rdr["IPLblShipLotNo"]) ? rdr["IPLblShipLotNo"].ToString() : null;
        //        worksheet.Cell(i, 9).Value = !Convert.IsDBNull(rdr["KitRepled"]) ? rdr["KitRepled"].ToString() : null;
        //        worksheet.Cell(i, 10).Value = !Convert.IsDBNull(rdr["notdone"]) ? rdr["notdone"].ToString() : null;
        //        worksheet.Cell(i, 11).Value = !Convert.IsDBNull(rdr["ReaYes"]) ? rdr["ReaYes"].ToString() : null;
        //        //worksheet.Cell(i, 12).Value = !Convert.IsDBNull(rdr["SFDATE"]) ? rdr["SFDATE"].ToString() : null;
        //        i++;

        //    }
        //    conn.Close();
        //    worksheet.Columns().AdjustToContents();
        //    return View(ImportExceltoDatatable(workbook));

        //}

        public IActionResult SubjectVisitReport()
        {
            string sql = "SELECT ROW_KEY, SITEID, SUBJID, BRTHDTC, SEX, AgeGroup, ICDTC, SCRNDTC, STATUS_INFO, PMCOMDATE, ADDDATE, ADDUSER, DATE_RAND, RANDBY, SFDATE FROM BIL_SUBJ WHERE STATUS_INFO = 'Randomized' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey")  + " ORDER BY SITEID   ";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Subject_Visit" + DateTime.Now.ToString("MM-dd-yyyy"));


            worksheet.Cell(1, 1).Value = "Site ID";
            worksheet.Cell(1, 2).Value = "Subject ID";
            worksheet.Cell(1, 3).Value = "Year of Birth";
            worksheet.Cell(1, 4).Value = "Sex";
            worksheet.Cell(1, 5).Value = "Status";
            worksheet.Cell(1, 6).Value = "Randomization Date";
            worksheet.Cell(1, 7).Value = "Visit 2";
            worksheet.Cell(1, 8).Value = "Visit 4";
            worksheet.Cell(1, 9).Value = "Visit 6";
            worksheet.Cell(1, 10).Value = "Visit 8";

            int i = 2;

            while (rdr.Read())
            {
                worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["SITEID"]) ? rdr["SITEID"].ToString() : null;
                worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["SUBJID"]) ? rdr["SUBJID"].ToString() : null;
                worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["BRTHDTC"]) ? rdr["BRTHDTC"].ToString() : null;
                worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["SEX"]) ? rdr["SEX"].ToString() : null;
                worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["STATUS_INFO"]) ? rdr["STATUS_INFO"].ToString() : null;
                worksheet.Cell(i, 6).Value = !Convert.IsDBNull(rdr["DATE_RAND"]) ? rdr["DATE_RAND"].ToString() : null;

                int ROWKEY = (int)rdr["ROW_KEY"];
                GetSubjVisitDet(worksheet.Row(i), ROWKEY);

                i++;

            }
            conn.Close();
            worksheet.Columns().AdjustToContents();
            return View(ImportExceltoDatatable2(workbook));

        }


        public void GetSubjVisitDet(IXLRow row, int rowkey)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            string sqlState;
            sqlState = "SELECT VISITID, REPLACE(UPPER(CONVERT(varchar, ADDDATE, 106)), ' ', '/') AS DATEASSGN " +
                       "FROM [BIL_VISITS] " +
                       "WHERE [ROW_KEY] = " + rowkey + " ORDER BY [ADDDATE]";


            SqlCommand cmd = new SqlCommand(sqlState, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string visitId = rdr["VISITID"].ToString();
                string dateAssgn = rdr["DATEASSGN"].ToString();

                if (!string.IsNullOrEmpty(visitId))
                {
                    int columnIndex;
                    if (int.TryParse(visitId.Replace("Visit", ""), out columnIndex))
                    {
                        // Assuming columnIndex starts from 7 (Visit 2 is the 7th column in your example)
                        // Update the ClosedXML row with visit details
                        try
                        {
                            if(columnIndex == 2)
                             row.Cell(6 + 1).Value = dateAssgn;
                            if(columnIndex == 4)
                                row.Cell(6 + 2).Value = dateAssgn;
                            if (columnIndex == 6)
                                row.Cell(6 + 3).Value = dateAssgn;
                            if (columnIndex == 8)
                                row.Cell(6 + 4).Value = dateAssgn;
                        }
                        catch (Exception ex)
                        {
                            // Handle the exception or log it for further investigation
                            // For example, you can log the values of columnIndex and visitId to identify the issue
                            Console.WriteLine($"Error updating cell for columnIndex {columnIndex}, visitId {visitId}: {ex.Message}");
                        }
                    }
                    else
                    {
                        // Handle the case where parsing the column index fails
                        // You may want to log a warning or handle this situation accordingly
                    }
                }


            }

            if (rdr != null) rdr.Close();
            if (conn != null) conn.Close();
        }


        public IActionResult SubjectVisitReportExcel()
        {
            //string SITEID = HttpContext.Session.GetString("sesCenter");
            //XLWorkbook wb = new XLWorkbook();
            //connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            //DataTable dt = CreateDataTable("SELECT SITEID AS 'Site ID', SUBJID AS 'Subject ID', VISIT AS 'Visit ID', KitNumber AS 'Kit Number', REPLACE(UPPER(CONVERT(varchar, [ASSIGNMENT_DATE], 106)), ' ', '/') AS 'Dispensed Date', ASSIGNED_BY AS 'Dispensed By', IPLblShipExpiryDate AS 'Expiry Date', IPLblShipLotNo AS 'Lot Number', KitRepled AS 'Replacement For', notdone AS 'Visit not Done', ReaYes AS 'Reason(IP dispensed, visit not done)' FROM BIL_IP_RANGE WHERE SITEID is not NULL AND SUBJID is not null ORDER BY SITEID, SUBJID", connectionString);
            ////DataTable dt = CreateDataTable("SELECT BIL_IP_RANGE.SITEID, BIL_IP_RANGE.KitNumber, BIL_IP_RANGE.KITSTAT, BIL_IP_RANGE.SENDBY, REPLACE(UPPER(CONVERT(varchar, BIL_IP_RANGE.SENDDATE, 106)), ' ', '/') AS 'SEND DATE', BIL_IP_RANGE.RECVDBY, REPLACE(UPPER(CONVERT(varchar, BIL_IP_RANGE.RECVDDATE, 106)), ' ', '/') AS 'RECVD DATE', BIL_IP_RANGE.ASSIGNED, BIL_IP_RANGE.ASSIGNMENT_DATE, BIL_IP_RANGE.ASSIGNED_BY, BIL_IP_RANGE.VISIT, BIL_SUBJ.SUBJID, BIL_IP_RANGE.KITCOMM FROM BIL_SUBJ RIGHT OUTER JOIN BIL_IP_RANGE ON BIL_SUBJ.ROW_KEY = BIL_IP_RANGE.ROW_KEY WHERE ((BIL_IP_RANGE.SITEID = '" + SITEID + "' OR '" + SITEID + "' = '(All)') AND ([SENDBY] IS NOT NULL)) ORDER BY BIL_IP_RANGE.SITEID, KitNumber", connectionString);

            //wb.Worksheets.Add(dt, "Subject Visit Report").Columns().AdjustToContents();
            //using (var stream = new MemoryStream())
            //{
            //    wb.SaveAs(stream);
            //    var content = stream.ToArray();

            //    return File(
            //        content,
            //        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            //        "Subject Visit Report" + DateTime.Now.Year + ".xlsx");
            //}
            string sql = "SELECT ROW_KEY, SITEID, SUBJID, BRTHDTC, SEX, AgeGroup, ICDTC, SCRNDTC, STATUS_INFO, PMCOMDATE, ADDDATE, ADDUSER, DATE_RAND, RANDBY, SFDATE FROM BIL_SUBJ WHERE STATUS_INFO = 'Randomized' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID   ";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Subject_Visit" + DateTime.Now.ToString("MM-dd-yyyy"));


            worksheet.Cell(1, 1).Value = "Site ID";
            worksheet.Cell(1, 2).Value = "Subject ID";
            worksheet.Cell(1, 3).Value = "Year of Birth";
            worksheet.Cell(1, 4).Value = "Sex";
            worksheet.Cell(1, 5).Value = "Status";
            worksheet.Cell(1, 6).Value = "Randomization Date";
            worksheet.Cell(1, 7).Value = "Visit 2";
            worksheet.Cell(1, 8).Value = "Visit 4";
            worksheet.Cell(1, 9).Value = "Visit 6";
            worksheet.Cell(1, 10).Value = "Visit 8";

            int i = 2;

            while (rdr.Read())
            {
                worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["SITEID"]) ? rdr["SITEID"].ToString() : null;
                worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["SUBJID"]) ? rdr["SUBJID"].ToString() : null;
                worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["BRTHDTC"]) ? rdr["BRTHDTC"].ToString() : null;
                worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["SEX"]) ? rdr["SEX"].ToString() : null;
                worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["STATUS_INFO"]) ? rdr["STATUS_INFO"].ToString() : null;
                worksheet.Cell(i, 6).Value = !Convert.IsDBNull(rdr["DATE_RAND"]) ? rdr["DATE_RAND"].ToString() : null;

                int ROWKEY = (int)rdr["ROW_KEY"];
                GetSubjVisitDet(worksheet.Row(i), ROWKEY);

                i++;

            }
            conn.Close();
            worksheet.Columns().AdjustToContents();
            worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 10)).Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
            worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 10)).Style.Font.FontColor = XLColor.White;
            worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 10)).Style.Font.Bold = true;
            worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 10)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //worksheet.SheetView.FreezeRows(1);
            //worksheet.SheetView.FreezeColumns(1);
            //worksheet.SheetView.FreezeColumns(2);
            worksheet.RangeUsed().SetAutoFilter();

            using (var stream = new MemoryStream())
            {
                workbook.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   "Subject Visit Report" + DateTime.Now.Year + ".xlsx");
            }



        }


        public IActionResult SiteIPReport()
        {
            string sql = "SELECT SITEID, KitNumber, KITTYPE, SENDBY, RECVDBY, RECVDDATE, KITSTAT, ASSIGNED, VISIT, SUBJID, KITCOMM, ASSIGNMENT_DATE, SENDDATE, IPLblShipExpiryDate, IPLblShipLotNo, ASSIGNED_BY FROM BIL_IP_RANGE WHERE (SENDBY IS NOT NULL) ORDER BY SITEID, KitNumber";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("SiteIP" + DateTime.Now.ToString("MM-dd-yyyy"));


            worksheet.Cell(1, 1).Value = "Site ID";
            worksheet.Cell(1, 2).Value = "Kit Number";
            worksheet.Cell(1, 3).Value = "Kit Type";
            worksheet.Cell(1, 4).Value = "Kit Status";
            worksheet.Cell(1, 5).Value = "Received By";
            worksheet.Cell(1, 6).Value = "Received Date";
            worksheet.Cell(1, 7).Value = "Sent By";
            worksheet.Cell(1, 8).Value = "Sent Date";
            worksheet.Cell(1, 9).Value = "Assigned";
            worksheet.Cell(1, 10).Value = "Subject ID";
            worksheet.Cell(1, 11).Value = "Visit ID";
            worksheet.Cell(1, 12).Value = "Assigment Date";
            worksheet.Cell(1, 13).Value = "Assigned By";
            worksheet.Cell(1, 14).Value = "Expiry Date";
            worksheet.Cell(1, 15).Value = "Lot NUmber";
            worksheet.Cell(1, 16).Value = "Comments";


            int i = 2;

            while (rdr.Read())
            {
                worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["SITEID"]) ? rdr["SITEID"].ToString() : null;
                worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["KitNumber"]) ? rdr["KitNumber"].ToString() : null;
                worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["KITTYPE"]) ? rdr["KITTYPE"].ToString() : null;
                worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["KITSTAT"]) ? rdr["KITSTAT"].ToString() : null;
                worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["RECVDBY"]) ? rdr["RECVDBY"].ToString() : null;
                worksheet.Cell(i, 6).Value = !Convert.IsDBNull(rdr["RECVDDATE"]) ? rdr["RECVDDATE"].ToString() : null;
                worksheet.Cell(i, 7).Value = !Convert.IsDBNull(rdr["SENDBY"]) ? rdr["SENDBY"].ToString() : null;
                worksheet.Cell(i, 8).Value = !Convert.IsDBNull(rdr["SENDDATE"]) ? rdr["SENDDATE"].ToString() : null;
                worksheet.Cell(i, 9).Value = !Convert.IsDBNull(rdr["ASSIGNED"]) ? rdr["ASSIGNED"].ToString() : null;
                worksheet.Cell(i, 10).Value = !Convert.IsDBNull(rdr["SUBJID"]) ? rdr["SUBJID"].ToString() : null;
                worksheet.Cell(i, 11).Value = !Convert.IsDBNull(rdr["VISIT"]) ? rdr["VISIT"].ToString() : null;
                worksheet.Cell(i, 12).Value = !Convert.IsDBNull(rdr["ASSIGNMENT_DATE"]) ? rdr["ASSIGNMENT_DATE"].ToString() : null;
                worksheet.Cell(i, 13).Value = !Convert.IsDBNull(rdr["ASSIGNED_BY"]) ? rdr["ASSIGNED_BY"].ToString() : null;
                worksheet.Cell(i, 14).Value = !Convert.IsDBNull(rdr["IPLblShipExpiryDate"]) ? rdr["IPLblShipExpiryDate"].ToString() : null;
                worksheet.Cell(i, 15).Value = !Convert.IsDBNull(rdr["IPLblShipLotNo"]) ? rdr["IPLblShipLotNo"].ToString() : null;
                worksheet.Cell(i, 16).Value = !Convert.IsDBNull(rdr["KITCOMM"]) ? rdr["KITCOMM"].ToString() : null;
                //worksheet.Cell(i, 10).Value = !Convert.IsDBNull(rdr["notdone"]) ? rdr["notdone"].ToString() : null;
                //worksheet.Cell(i, 11).Value = !Convert.IsDBNull(rdr["ReaYes"]) ? rdr["ReaYes"].ToString() : null;
                //worksheet.Cell(i, 12).Value = !Convert.IsDBNull(rdr["SFDATE"]) ? rdr["SFDATE"].ToString() : null;
                i++;

            }
            conn.Close();
            worksheet.Columns().AdjustToContents();
            return View(ImportExceltoDatatable(workbook));

        }

        public IActionResult SiteIPReportExcel()
        {

            XLWorkbook wb = new XLWorkbook();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DataTable dt = CreateDataTable("SELECT SITEID AS 'Site ID', KitNumber AS 'Kit Number', KITTYPE AS 'Kit Type', KITSTAT AS 'Kit Status', SENDBY AS 'Send By', SENDDATE AS 'Sent Date', RECVDBY AS 'Received By', RECVDDATE AS 'Received Date', ASSIGNED AS 'Assigned', VISIT AS 'Visit ID', SUBJID AS 'Subject ID', ASSIGNMENT_DATE AS 'Dispensed Date',  ASSIGNED_BY AS 'Assigned By', KITCOMM AS 'Comments', IPLblShipExpiryDate AS 'Expiry Date', IPLblShipLotNo AS 'Lot Number' FROM BIL_IP_RANGE WHERE (SENDBY IS NOT NULL)  ORDER BY SITEID, KitNumber", connectionString);
            wb.Worksheets.Add(dt, "Site IP Report").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Site IP Report" + DateTime.Now.Year + ".xlsx");
            }

        }

        public IActionResult DepotIPReport()
        {
            string sql = "SELECT KITKEY, KitNumber, KITSTAT, IPLblShipExpiryDate, IPLblShipLotNo, KITCOMM FROM BIL_IP_RANGE WHERE (SITEID IS  NULL)  ORDER BY  KitNumber";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("IPDepot" + DateTime.Now.ToString("MM-dd-yyyy"));


            worksheet.Cell(1, 1).Value = "Kit Key";
            worksheet.Cell(1, 2).Value = "Kit Number";
            worksheet.Cell(1, 3).Value = "Kit Status";
            worksheet.Cell(1, 4).Value = "Expiry Date";
            worksheet.Cell(1, 5).Value = "Lot NUmber";
            worksheet.Cell(1, 6).Value = "Comments";


            int i = 2;

            while (rdr.Read())
            {
                worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["KITKEY"]) ? rdr["KITKEY"].ToString() : null;
                worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["KitNumber"]) ? rdr["KitNumber"].ToString() : null;
                worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["KITSTAT"]) ? rdr["KITSTAT"].ToString() : null;
                worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["IPLblShipExpiryDate"]) ? rdr["IPLblShipExpiryDate"].ToString() : null;
                worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["IPLblShipLotNo"]) ? rdr["IPLblShipLotNo"].ToString() : null;
                worksheet.Cell(i, 6).Value = !Convert.IsDBNull(rdr["KITCOMM"]) ? rdr["KITCOMM"].ToString() : null;
                //worksheet.Cell(i, 10).Value = !Convert.IsDBNull(rdr["notdone"]) ? rdr["notdone"].ToString() : null;
                //worksheet.Cell(i, 11).Value = !Convert.IsDBNull(rdr["ReaYes"]) ? rdr["ReaYes"].ToString() : null;
                //worksheet.Cell(i, 12).Value = !Convert.IsDBNull(rdr["SFDATE"]) ? rdr["SFDATE"].ToString() : null;
                i++;

            }
            conn.Close();
            worksheet.Columns().AdjustToContents();
            return View(ImportExceltoDatatable(workbook));

        }

        public IActionResult DepotIPReportExcel()
        {

            XLWorkbook wb = new XLWorkbook();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DataTable dt = CreateDataTable("SELECT KITKEY AS 'Kit Key', KitNumber AS 'Kit Number', KITSTAT AS 'Kit Status',  KITCOMM AS 'Comments', IPLblShipExpiryDate AS 'Expiry Date', IPLblShipLotNo AS 'Lot Number' FROM BIL_IP_RANGE WHERE (SITEID IS  NULL) ORDER BY  KitNumber", connectionString);
            wb.Worksheets.Add(dt, "IPDepot Report").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "IP Depot Report" + DateTime.Now.Year + ".xlsx");
            }

        }


        public IActionResult ViewIPDepot()
        {
            List<DepotInv> Inv = new List<DepotInv>();
            string sql = "SELECT KITKEY, KitNumber, KITSTAT, KITSET, SITEID FROM BIL_IP_RANGE WHERE (SITEID IS  NULL) OR (KITSTAT = 'Shipped')  ORDER BY  KitNumber";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                DepotInv temp = new DepotInv();

                temp.KitNumber = rdr["KitNumber"].ToString();
                temp.KITSET = rdr["KITSET"].ToString(); ;
                temp.SITEID = rdr["SITEID"].ToString();
                temp.KITSTAT = rdr["KITSTAT"].ToString();
                temp.KITKEY = (int)rdr["KITKEY"];
                Inv.Add(temp);
            }
            
            conn.Close();
            return View(Inv);

        }


        public IActionResult Visits(int ROW_KEY, string SUBJID)
        {

            //int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

            // string sql = "SELECT VISITKEY, VISITID, SITEID, SUBJID, ROW_KEY, KitNumber, ADDDATE, ADDUSER FROM BIL_VISITS WHERE SUBJID = '" + SUBJID + "' ORDER BY KitNumber";
            string sql = "SELECT [VISITKEY], ROW_KEY, VISITID, SITEID, SUBJID, VISITCOMM, VisitDate, ELIGIP, ReaYes, ADDUSER  FROM BIL_VISITS WHERE (ROW_KEY = " + ROW_KEY + " ) ORDER BY VISITKEY";

            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            var model = new List<Visits>();
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var subjects = new Visits();
                    subjects.VISITID = rdr["VISITID"].ToString();
                    subjects.SITEID = rdr["SITEID"].ToString();
                    subjects.SUBJID = rdr["SUBJID"].ToString();
                    subjects.VISITCOMM = rdr["VISITCOMM"].ToString();
                    subjects.ADDUSER = rdr["ADDUSER"].ToString();
                    subjects.ROW_KEY = (int)rdr["ROW_KEY"];
                    subjects.VISITKEY = (int)rdr["VISITKEY"];
                    subjects.VisitDate = rdr["VisitDate"].ToString();
                    subjects.ELIGIP = rdr["ELIGIP"].ToString();
                    subjects.ReaYes = rdr["ReaYes"].ToString();

                    model.Add(subjects);
                    TempData["SUBJID"] = SUBJID;
             
                }

            }
            con.Close();
            TempData["SUBJID"] = SUBJID;
            //  TempData.Keep("SUBJID");
            return View(model);

        }


    }

}



