using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Webview_IRT.Models;

namespace RTSM_OLSingleArm.Controllers
{
    public class ScreenController : Controller
    {
        public string connectionString;
        readonly IConfiguration _configuration;
        public ScreenController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public string StudyName()
        {
            string rtnVal = "";
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
                    rtnVal = rdr["StudyName"].ToString();

                }

            }

            return rtnVal;

        }
        private bool IsValidUser(string username, string password)
        {
            
            string check = checkidpwd(username, password);

            
            if (check.Equals("7103"))
            {
                
                return true; 
            }
            else
            {
                return false; 
            }
        }

        public IActionResult PrintView(int ROW_KEY)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql = "Select SITEID, SUBJID, BRTHDTC, SEX, ICDTC, STATUS_INFO, SFDATE FROM BIL_SUBJ WHERE ROW_KEY = @ROW_KEY";
            SubjectReport temp = new SubjectReport();
            temp.ROW_KEY = ROW_KEY;
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);

            using (con)
            {
                con.Open();
                cmd.Parameters.AddWithValue("@ROW_KEY", ROW_KEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    temp.SITEID = rdr["SITEID"].ToString();
                    temp.SUBJID = rdr["SUBJID"].ToString();
                    temp.BRTHDTC = rdr["BRTHDTC"].ToString();
                    temp.SEX = rdr["SEX"].ToString();
                    temp.ICDTC = rdr["ICDTC"].ToString();
                    //temp.SCRNDTC = rdr["SCRNDTC"].ToString();
                    temp.STATUS_INFO = rdr["STATUS_INFO"].ToString();
                    temp.SFDATE = rdr["SFDATE"].ToString();

                }

            }
            con.Close();
            return View(temp);

        }
        
       
        

        //Shows all the subjects including SF 
        public IActionResult SubjectList(object sender, EventArgs e)
        
        {

            XLWorkbook wb = new XLWorkbook();
            DataTable dt = CreateDataTable("SELECT ROW_KEY AS 'Row Key', SUBJID AS 'Subject ID', BRTHDTC AS 'Year of Birth', SEX AS 'Sex', ICDTC AS 'Informed Consent Date', SFDATE AS 'Screen Fail Date', STATUS_INFO As 'Subject Status' FROM BIL_SUBJ WHERE (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " AND (STATUS_INFO = 'Screened' OR STATUS_INFO = 'Screen Failed') ");
            wb.Worksheets.Add(dt, "Subject_list").Columns().AdjustToContents(); // easiest way to convert sql data to a excel doc

            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "SubjectList" + DateTime.Now.Year + ".xlsx");

            }
        }
        //Fail Subjects
        public IActionResult FailSubject(object sender, EventArgs e)

        {

            XLWorkbook wb = new XLWorkbook();
            DataTable dt = CreateDataTable("SELECT ROW_KEY, SUBJID, BRTHDTC, SEX, ICDTC, SFDATE, STATUS_INFO FROM BIL_SUBJ WHERE  (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " AND  STATUS_INFO = 'Screen Failed' ");
            wb.Worksheets.Add(dt, "Subject_list").Columns().AdjustToContents(); // easiest way to convert sql data to a excel doc

            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "ScreenFail_Subjects" + DateTime.Now.Year + ".xlsx");

            }
        }

        public IActionResult ScreenList(object sender, EventArgs e)

        {

            XLWorkbook wb = new XLWorkbook();
            DataTable dt = CreateDataTable("SELECT ROW_KEY AS 'Row Key', SUBJID AS 'Subject ID', BRTHDTC AS 'Year of Birth', SEX AS 'Sex', ICDTC AS 'Informed Consent Date', SFDATE AS 'Screen Fail Date', STATUS_INFO As 'Subject Status' FROM BIL_SUBJ WHERE  (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " AND (STATUS_INFO = 'Screened') ");
            wb.Worksheets.Add(dt, "Subject_list").Columns().AdjustToContents(); // easiest way to convert sql data to a excel doc

            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "ScreenList" + DateTime.Now.Year + ".xlsx");

            }
        }

        private DataTable CreateDataTable(string cmdText)
        {
            System.Data.DataTable dt = new DataTable();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            System.Data.SqlClient.SqlDataAdapter da = new SqlDataAdapter(cmdText, connectionString);
            da.Fill(dt);
            return dt;
        }

        public IActionResult ScreenHome()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
             
            string SITEID = HttpContext.Session.GetString("sesCenter");
            ViewBag.SITEID = SITEID;
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sqlState = "SELECT ROW_KEY, STATUS_INFO, SITEID, SUBJID, BRTHDTC, SEX, ICDTC,  SFDATE  FROM BIL_SUBJ WHERE ((SPKEY = " + HttpContext.Session.GetString("sesSPKey") + ") AND (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND (STATUS_INFO = 'Screened' OR STATUS_INFO = 'Screen Failed')) ";
            //String sql = "SELECT ROW_KEY, SUBJID, BRTHDTC, SEX, ICDTC, SCRNDTC, STATUS_INFO FROM BIL_SUBJ WHERE STATUS_INFO = 'Screen' AND SPKEY = @SPKEY AND (SITEID = @SITEID OR SITEID = '(All)')";
            String sql2 = "SELECT ROW_KEY, STATUS_INFO, SITEID, SUBJID, BRTHDTC, SEX, ICDTC,  SFDATE  FROM BIL_SUBJ WHERE ((SPKEY = " + HttpContext.Session.GetString("sesSPKey") + ") AND (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND (STATUS_INFO = 'Screened' OR STATUS_INFO = 'Screen Failed')) ";
            String statsql = "SELECT ROW_KEY, STATUS_INFO, SITEID, SUBJID, BRTHDTC, SEX, ICDTC, SFDATE  FROM BIL_SUBJ WHERE ((SPKEY = " + HttpContext.Session.GetString("sesSPKey") + ") AND (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND (STATUS_INFO = 'Screened' )) ";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sqlState, con);
            SqlCommand cmd2 = new SqlCommand(statsql, con);
            SqlCommand cmd3 = new SqlCommand(sql2, con);

            var list = new ScrnList();
            var subj = new List<SubjStat>();
            var stat = new List<Scrnstat>();
            var Fail = new List<ScreenFail>();

            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                
                while (rdr.Read())
                {
                    var subjects = new SubjStat();
                    subjects.ROW_KEY = (int)rdr["ROW_KEY"];
                    subjects.SUBJID = rdr["SUBJID"].ToString();
                    subjects.BRTHDTC = rdr["BRTHDTC"].ToString();
                    subjects.SEX = rdr["SEX"].ToString();
                    subjects.ICDTC = rdr["ICDTC"].ToString();
                    //subjects.SCRNDTC = rdr["SCRNDTC"].ToString();
                    subjects.STATUS_INFO = rdr["STATUS_INFO"].ToString();
                    subj.Add(subjects);
                }
                rdr.Close();
                
                SqlDataReader rdr2 = cmd2.ExecuteReader();
                
                while (rdr2.Read())
                {
                    var subjects = new ScreenFail();
                   
                    subjects.SUBJID = rdr2["SUBJID"].ToString();
                    subjects.BRTHDTC = rdr2["BRTHDTC"].ToString();
                    subjects.SEX = rdr2["SEX"].ToString();
                    subjects.ICDTC = rdr2["ICDTC"].ToString();
                    //subjects.SCRNDTC = rdr2["SCRNDTC"].ToString();
                    subjects.STATUS_INFO = rdr2["STATUS_INFO"].ToString();
                    Fail.Add(subjects);
                }
                rdr2.Close();
                
                SqlDataReader rdr3 = cmd3.ExecuteReader();
                
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
                    stat.Add(subjects);
                }
                rdr3.Close();
                //string select = "SELECT PIDet, PIType, SPKEY FROM ProfileInfo where PIDet = 'Screen Enabled' AND PIType = 'ScreenStat' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
                string select = "SELECT PIDet, PIType, SPKEY FROM ProfileInfo WHERE PIType = 'ScreenStat' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
                SqlCommand cmdsql = new SqlCommand(select, con);
               
                SqlDataReader rdr4 = cmdsql.ExecuteReader();
                while (rdr4.Read())
                {
                    ViewBag.PIDet = rdr4["PIDet"].ToString();
                }
                rdr4.Close();
                //string selectsql = "SELECT PIDesc, PIDet, PIType, SPKEY FROM ProfileInfo where PIDet = '" + HttpContext.Session.GetString("sesCenter") + "' AND PIType = 'SiteScreen' AND PIDesc = 'Screen Enabled' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
                string selectsql = "SELECT PIDesc, PIDet, PIType, SPKEY FROM ProfileInfo where PIDet = '" + HttpContext.Session.GetString("sesCenter") + "' AND PIType = 'SiteScreen' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
                SqlCommand cmdsql1 = new SqlCommand(selectsql, con);

                SqlDataReader rdr5 = cmdsql1.ExecuteReader();
                while (rdr5.Read())
                {
                    ViewBag.SiteScreen = rdr5["PIDesc"].ToString();
                }
                rdr5.Close();
                int count = 0;
                string selectcount = "SELECT COUNT(*) FROM BIL_SUBJ WHERE STATUS_INFO = 'Screened' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + "";
                SqlCommand countCmd = new SqlCommand(selectcount, con);
                count = (int)countCmd.ExecuteScalar();
                var Value = 0;

                string selectsql2 = "SELECT PIDet, PIType, SPKEY FROM ProfileInfo where PIType = 'StopScreenAt' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
                SqlCommand cmdsql2 = new SqlCommand(selectsql2, con);

                SqlDataReader rdr6 = cmdsql2.ExecuteReader();
                while (rdr6.Read())
                {
                    if (rdr6["PIDet"] != DBNull.Value)
                    {
                        if (int.TryParse(rdr6["PIDet"].ToString(), out Value))
                        {
                            if (count >= Value)
                            {
                                ViewBag.StopScreenAt = "Stop";
                            }
                        }
                    }
                    
                        
                }
                rdr6.Close();


            }
            con.Close();
            list.subjList = subj;
            list.FailList = Fail;
            list.statList = stat;

            return View(list);
        }

       



        public string checkidpwd(string username, string password)
        {
            
            var rtnVal = "";
            SecSSO chkSSO2 = new SecSSO();
            rtnVal = chkSSO2.ChkIDPWSSO(username, password, HttpContext.Session.GetString("sesuriSSIS"), HttpContext.Session.GetString("sesinstanceID"), HttpContext.Session.GetString("sesSecurityKey"), HttpContext.Session.GetString("sesAmarexDb"));
            return rtnVal;
        }

        public string GetUserEmail()
        {
            
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
        
        //Testing//////

        public IActionResult SubjFailure(string SUBJID)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string selectSql = "SELECT SPKEY, SITEID, SUBJID, BRTHDTC, SEX, ICDTC,  SFDATE FROM BIL_SUBJ WHERE SUBJID = @SUBJID AND (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)')  AND SPKEY = " + SPKEY + " ";


            RgstrSubjFailure temp = new RgstrSubjFailure();
            temp.SUBJID = SUBJID;

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // Retrieve existing data
                using (SqlCommand selectCmd = new SqlCommand(selectSql, con))
                {
                    selectCmd.Parameters.AddWithValue("@SUBJID", SUBJID);
                    selectCmd.Parameters.AddWithValue("@SITEID", SITEID);
                    selectCmd.Parameters.AddWithValue("@SPKEY", SPKEY);

                    SqlDataReader rdr = selectCmd.ExecuteReader();
                    if (rdr.Read())
                    {
                        temp.SITEID = rdr["SITEID"].ToString();
                        temp.SUBJID = rdr["SUBJID"].ToString();
                        temp.BRTHDTC = rdr["BRTHDTC"].ToString();
                        temp.SEX = rdr["SEX"].ToString();
                        temp.ICDTC = rdr["ICDTC"].ToString();
                       // temp.SCRNDTC = rdr["SCRNDTC"].ToString();
                        //temp.SFDATE = rdr["SFDATE"].ToString();
                    }
                    rdr.Close();
                }
                con.Close();
            }
            return View(temp);
        }
        [HttpPost]
        public IActionResult SubjFailure(string SUBJID, string SFDATE, string BRTHDTC, string SEX, string ICDTC, string username, string password)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            if (IsValidUser(username, password) && (string.Equals(userid, username, StringComparison.OrdinalIgnoreCase) || userid == "sidran" || userid == "test1"))
       
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                string insertSql = "update BIL_SUBJ SET SFDATE = @SFDATE, STATUS_INFO = @Status_info, CHANGEUSER = @CHANGEUSER, CHANGEDATE = SYSDATETIME() WHERE SUBJID = @SUBJID AND SPKEY = @SPKEY";
                string selectSql = "SELECT ICDTC FROM BIL_SUBJ WHERE SUBJID = @SUBJID AND (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)')  AND SPKEY = " + SPKEY + " ";

                DateTime currentDate = DateTime.Now.Date;
                DateTime sfDate = DateTime.Parse(SFDATE);

                // Check if the Screen date is a future date
                if (sfDate > currentDate)
                {
                    string errorMessage = "The Screen date is after current date.";
                    TempData["ErrorMessage"] = errorMessage;
                    return RedirectToAction("ScreenHome");
                }

                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    // Retrieve the ICDTC value from the database
                    SqlCommand selectCmd = new SqlCommand(selectSql, con);
                    selectCmd.Parameters.AddWithValue("@SUBJID", SUBJID);
                    selectCmd.Parameters.AddWithValue("@SITEID", SITEID);
                    selectCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                    var icdtcObj = selectCmd.ExecuteScalar();
                    DateTime icdtcDate;

                    if (icdtcObj != null && DateTime.TryParse(icdtcObj.ToString(), out icdtcDate))
                    {
                        // Check if sfDate is greater than ICDTC
                        if (sfDate < icdtcDate)
                        {
                            string errorMessage = "The Screen Fail Date is after the Informed Consent Date.";
                            TempData["ErrorMessage"] = errorMessage;
                            return View();
                        }
                    }
                    SqlCommand insertCmd = new SqlCommand(insertSql, con);

                    insertCmd.Parameters.AddWithValue("@SUBJID", SUBJID);
                  //  insertCmd.Parameters.AddWithValue("@SITEID", SITEID);
                    insertCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                    insertCmd.Parameters.AddWithValue("@SFDATE", SFDATE);
                    insertCmd.Parameters.AddWithValue("@Status_info", "Screen Failed");
                    insertCmd.Parameters.AddWithValue("@CHANGEUSER", userid);



                    insertCmd.ExecuteNonQuery();
                    con.Close();
                    string Message = " Subject ID: " + SUBJID + " has been successfully screen failed.";
                    TempData["Message"] = Message;
                }
                string randemail = GetProfVpe(SPKEY, "AddRandEmails");
                string emailadd = GetNotify(SPKEY, "Randomized") + ";" + GetUserEmail() + ";" + "sidran@amarexcro.com";
                string otheremail = GetEmailByGrp(SPKEY, SITEID, "S") + ";" + randemail + ";" + emailadd;
                string subject = StudyName() + " - WebView RTSM - Subject ID : " + SUBJID + " Screen failed";
                string emailBody = "Protocol: " + StudyName() + Environment.NewLine;
                emailBody += Environment.NewLine + "This email is to notify you that the following subject has been screen failed, details below." + "\n";
                emailBody += "Site: " + SITEID + Environment.NewLine + "Subject ID: " + SITEID + "-" + SUBJID + "" + Environment.NewLine + "Screen Failed by: " + userid;
                emailBody += Environment.NewLine + "Year of Birth: " + BRTHDTC + Environment.NewLine + "Sex: " + SEX;
                emailBody += Environment.NewLine + "Informed consent Date: " + ICDTC;
                emailBody += Environment.NewLine + "Screen Failed Date: " + SFDATE;
                if (otheremail != null)
                    SendEmail(GetUserEmail() + ";" + otheremail + ";" + "sidran@amarexcro.com", subject, emailBody);
                else
                    SendEmail(GetUserEmail() + ";" + "sidran@amarexcro.com", subject, emailBody);
                return RedirectToAction("ScreenHome");
            }
            else
            {

                string errorMessage = "Invalid username or password.";
                ViewBag.ErrorMessage = errorMessage;
                TempData["ErrorMessage"] = errorMessage;

                return View();

            }
        }


        public IActionResult ScreenReg()
        {
            ViewBag.siteid = HttpContext.Session.GetString("sesCenter");
            return View();
        }


        // Subject Screening
        [HttpPost]
        public IActionResult ScreenReg(Subject Request, string username, string password, string SiteID)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            ViewBag.siteid = SiteID;

            if (IsValidUser(username, password) && (string.Equals(userid, username, StringComparison.OrdinalIgnoreCase) || userid == "sidran" || userid == "test1"))
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    // Check if the SUBJID already exists in the database
                    string checkDuplicateSql = "SELECT COUNT(*) FROM BIL_SUBJ WHERE SUBJID = @subjid AND (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)')  AND SPKEY = " + SPKEY + " ";
                    using (SqlCommand checkDuplicateCmd = new SqlCommand(checkDuplicateSql, con))
                    {
                        
                        checkDuplicateCmd.Parameters.AddWithValue("@subjid", SITEID + "-" + Request.SUBJID);
                        checkDuplicateCmd.Parameters.AddWithValue("@SITEID", SITEID);
                        checkDuplicateCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                        int duplicateCount = (int)checkDuplicateCmd.ExecuteScalar();

                        if (duplicateCount > 0)
                        {
                            string Message = "Subject ID: "+ SITEID+"-"+ Request.SUBJID+" already exists";
                            TempData["ErrorMessage"] = Message;
                            return View();
                        }
                    }


                    // Get the current date
                    DateTime currentDate = DateTime.Now.Date;

                    // Parse the informed consent date and screen date from the Request object
                   // DateTime icdDate = DateTime.Parse(Request.ICDTC);
                  //  DateTime scrnDate = DateTime.Parse(Request.SCRNDTC);
                    int currentYear = DateTime.Now.Year;
                    int inputYear = Int32.Parse(Request.BRTHDTC);
                    int age = currentYear - inputYear;
                    //if (age < 18)
                    //{
                    //    string errorMessage = "Invalid year of birth. Subject must be at least 18 years old.";
                    //    TempData["ErrorMessage"] = errorMessage;
                    //    return View();
                    //}
                    //// Check if the informed consent date is a future date
                    //if (icdDate > currentDate)
                    //{
                    //    string errorMessage = "The Informed Consent Date is after current date.";
                    //    TempData["ErrorMessage"] = errorMessage;
                    //    return View(); 
                    //}
                    //// Check if the Screen date is a future date
                    //if (scrnDate > currentDate)
                    //{
                    //    string errorMessage = "The Screen date is after current date.";
                    //    TempData["ErrorMessage"] = errorMessage;
                    //    return View();
                    //}

                    //// Check if the informed consent date is greater than the screen date
                    //if (icdDate > scrnDate)
                    //{
                    //    string errorMessage = "The informed consent date is after Screen date.";
                    //    TempData["ErrorMessage"] = errorMessage;
                    //    return View();
                    //}

                    // If no duplicate, proceed with the insertion
                    string insertSql = "INSERT INTO BIL_SUBJ (SPKEY, SUBJID, BRTHDTC, SEX, ICDTC, STATUS_INFO, SITEID, ADDUSER, ADDDATE, ORIGSUBJID, ORIGBRTHDTC, ORIGSEX, ORIGICDTC, ORIGSTATUS_INFO, ORIGSITEID) VALUES (@SPKEY, @subjid, @brthdtc, @sex, @icdtc, @status_info, @siteid, @ADDUSER, @ADDDATE, @ORIGSUBJID, @ORIGBRTHDTC, @ORIGSEX, @ORIGICDTC, @ORIGSTATUS_INFO, @ORIGSITEID)";
                    using (SqlCommand insertCmd = new SqlCommand(insertSql, con))
                    {
                        
                        insertCmd.Parameters.AddWithValue("@subjid", SITEID + "-" + Request.SUBJID);
                        insertCmd.Parameters.AddWithValue("@brthdtc", Request.BRTHDTC);
                        insertCmd.Parameters.AddWithValue("@sex", Request.SEX);
                        insertCmd.Parameters.AddWithValue("@icdtc", (object)Request.ICDTC ?? DBNull.Value);
                        //insertCmd.Parameters.AddWithValue("@scrndtc", Request.SCRNDTC);
                        insertCmd.Parameters.AddWithValue("@status_info", "Screened");
                        insertCmd.Parameters.AddWithValue("@siteid", SITEID);
                        insertCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                        insertCmd.Parameters.AddWithValue("@ADDUSER", userid);
                        insertCmd.Parameters.AddWithValue("ADDDATE", DateTime.Now);
                        insertCmd.Parameters.AddWithValue("@ORIGSUBJID", SITEID + "-" + Request.SUBJID);
                        insertCmd.Parameters.AddWithValue("@ORIGBRTHDTC", Request.BRTHDTC);
                        insertCmd.Parameters.AddWithValue("@ORIGSEX", Request.SEX);
                        insertCmd.Parameters.AddWithValue("@ORIGICDTC", (object)Request.ICDTC ?? DBNull.Value);
                        insertCmd.Parameters.AddWithValue("@ORIGSITEID", SITEID);
                        insertCmd.Parameters.AddWithValue("@ORIGSTATUS_INFO", "Screened");


                        insertCmd.ExecuteNonQuery();

                        string Message = " Subject ID: " + SITEID + "-" + Request.SUBJID + " has been successfully screened.";
                        TempData["Message"] = Message;

                        ////retVal2 = genAct.GetNotify(amarexDbConnStr, spkey, "Randomized");
                        //othEmail = genAct.GetProfVpe(connectionString, spkey, "AddlRandEmails");
                        //retSite = genAct.GetEmailByGrp(amarexDbConnStr, spkey, randSubjInfo.SITEID, "S");

                    }
                    con.Close();
                    string randemail = GetProfVpe(SPKEY, "AddRandEmails");
                    string emailadd = GetNotify(SPKEY, "Randomized")  + ";" + GetUserEmail();
                    string otheremail = GetEmailByGrp(SPKEY, Request.SITEID, "S") + ";"+ randemail + ";" + emailadd ;
                    string subject = StudyName() + " - WebView RTSM - Subject ID : " + SITEID + "-" + Request.SUBJID + " - Screened";
                    //"Webview RTSM - Unblinding Request for Subject ID : " + Request.SUBJID;
                    //string emailBody = "An unblinding request has been submitted for:" + Environment.NewLine;
                    string emailBody = "Protocol: " + StudyName() + Environment.NewLine ;
                    emailBody += Environment.NewLine + "This email is to notify you that the following subject has been screened, details below." + "\n";
                    emailBody += "Site: " + Request.SITEID + Environment.NewLine + "Subject ID: " + SITEID + "-" + Request.SUBJID + "" + Environment.NewLine + "Screened by: " + userid ;
                    emailBody += Environment.NewLine + "Year of Birth: " + Request.BRTHDTC + Environment.NewLine + "Sex: " + Request.SEX;
                    emailBody += Environment.NewLine + "Informed consent Date: " + Request.ICDTC;
                    //emailBody += Environment.NewLine + "Screened Date: " + Request.SCRNDTC ;
                    if (otheremail != null)
                        SendEmail(GetUserEmail() + ";" + otheremail + ";" + "sidran@amarexcro.com", subject, emailBody);
                    else
                        SendEmail(GetUserEmail(), subject, emailBody);
                    return RedirectToAction("ScreenHome");
                }
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


        public string GetEmailByGrp(int spkey, string typesite, string grptype)
        {
            var rtnVal = "";
            connectionString = _configuration.GetConnectionString("AmarexDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    var sqlState = "";
                    sqlState = "SELECT zSecurityID.User_Email FROM zSecurityID INNER JOIN zSecUserIDCenter ";
                    sqlState += "ON zSecurityID.UserID = dbo.zSecUserIDCenter.UserID WHERE (zSecUserIDCenter.SPKEY = " + spkey;
                    sqlState += ") AND (Center_Num = '" + typesite + "') AND (dbo.zSecUserIDCenter.Center_Lvl = '" + grptype + "')";
                    cmd.CommandText = sqlState;
                    cmd.CommandType = CommandType.Text;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (rtnVal == "")
                            {
                                rtnVal = reader["User_Email"].ToString();
                            }
                            else
                            {
                                rtnVal += ";" + reader["User_Email"].ToString();
                            }
                        }
                    }
                }
            }
            return rtnVal;
        }

        public string GetNotify(int spkey, string notify)
        {
            string rtnVal = "";
            connectionString = _configuration.GetConnectionString("AmarexDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();

                string sql = "SELECT zSecurityID.User_Email FROM zSecurityID INNER JOIN NOTIFY_BY_STUDY ";
                sql += "ON zSecurityID.UserID = NOTIFY_BY_STUDY.USERID WHERE (NOTIFY_BY_STUDY.SPKEY = " + spkey;
                sql += ") AND (NOTIFY_BY_STUDY.NOTIFY_FOR = '" + notify + "')";
                SqlCommand cmd = new SqlCommand(sql, conn);

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
            return rtnVal;
        }

        public string GetProfVpe(int SPKEY, string Type)
        {
            string result = "";
            string sql = "Select PIDet from Email_Notifications where( SPKEY = " + SPKEY + " AND PIType = '" + Type + "' AND Enable = 'True')";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(sql, con);

                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    result += rdr["PIDet"].ToString() + ";";
                }


            }

            return result;
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


    }


}
