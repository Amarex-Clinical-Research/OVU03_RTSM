using AppUidAuth;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Webview_IRT.Models;

namespace RTSM_OLSingleArm.Controllers
{
    public class SubsequentDispensationController : Controller
    {
        public string connectionString;
        readonly IConfiguration _configuration;
        public SubsequentDispensationController(IConfiguration configuration)
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
        public IActionResult SubsequentDispensationHome()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            ViewBag.SITEID = SITEID;
            ViewBag.UserID = userid;

            string sql = "SELECT SPKEY, SITEID, SUBJID, ROW_KEY, ARM FROM BIL_SUBJ WHERE (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " AND STATUS_INFO = 'Randomized' ORDER BY SUBJID, ROW_KEY";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            var model = new List<Subject>();
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var subjects = new Subject();
                    subjects.SPKEY = (int)rdr["SPKEY"];
                   
                    subjects.SITEID = rdr["SITEID"].ToString();
                    subjects.SUBJID = rdr["SUBJID"].ToString();
                    subjects.ROW_KEY = (int)rdr["ROW_KEY"];
                    subjects.ARM = rdr["ARM"].ToString();
                    model.Add(subjects);
                }

            }
            con.Close();
            return View(model);

        }
        //public IActionResult SubjectVisits(int ROW_KEY, string SUBJID, int SPKEY)
        //{
            
        //    //int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
        //    string userid = HttpContext.Session.GetString("suserid");
        //    string SITEID = HttpContext.Session.GetString("sesCenter");
        //    connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

        //   // string sql = "SELECT VISITKEY, VISITID, SITEID, SUBJID, ROW_KEY, KitNumber, ADDDATE, ADDUSER FROM BIL_VISITS WHERE SUBJID = '" + SUBJID + "' ORDER BY KitNumber";
        //   string sql = "SELECT [VISIT], KitNumber, REPLACE(UPPER(CONVERT(varchar, [ASSIGNMENT_DATE], 106)), ' ', '/') AS 'Dispense', ASSIGNED_BY , IPLblShipExpiryDate, IPLblShipLotNo, KitRepled, SITEID, SUBJID,ROW_KEY FROM BIL_IP_RANGE WHERE (ROW_KEY = " + ROW_KEY + " AND SUBJID = '" + SUBJID + "') ORDER BY KitNumber";
           
        //    SqlConnection con = new SqlConnection(connectionString);
        //    SqlCommand cmd = new SqlCommand(sql, con);
        //    var model = new List<IPStatus>();
        //    using (con)
        //    {
        //        con.Open();
        //        SqlDataReader rdr = cmd.ExecuteReader();
        //        while (rdr.Read())
        //        {
        //            var subjects = new IPStatus();
        //            subjects.KitNumber = rdr["KitNumber"].ToString();
        //            subjects.VISIT = rdr["VISIT"].ToString();
        //            subjects.SITEID = rdr["SITEID"].ToString();
        //            subjects.SUBJID = rdr["SUBJID"].ToString();
        //            subjects.Dispense = rdr["Dispense"].ToString();
        //            subjects.ASSIGNED_BY = rdr["ASSIGNED_BY"].ToString();
        //            subjects.ROW_KEY = (int)rdr["ROW_KEY"];
        //            subjects.IPLblShipExpiryDate = rdr["IPLblShipExpiryDate"].ToString();
        //            subjects.IPLblShipLotNo = rdr["IPLblShipLotNo"].ToString();
        //            subjects.KitRepled = rdr["KitRepled"].ToString();

        //            model.Add(subjects);
        //            TempData["SUBJID"] = SUBJID;
        //            ViewBag.SPKEY = SPKEY;
        //        }

        //    }
        //    con.Close();
        //    TempData["SUBJID"] = SUBJID;
        //  //  TempData.Keep("SUBJID");
        //    return View(model);

        //}

        public IActionResult SubDispenseIP(int ROW_KEY, string SUBJID, string SITEID, string SPKEY )
        {
            ViewBag.SITEID = SITEID;
            ViewBag.SUBJID = SUBJID;
          //  return View();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

            string sql2 = "SELECT VISITKEY, VISITID, SITEID, SUBJID, ROW_KEY, ADDDATE, ADDUSER, ELIGIP FROM BIL_VISITS WHERE SUBJID = '" + SUBJID + "' AND ELIGIP = 'No' ORDER BY VISITKEY";
            string sql = "SELECT [VISIT], KitNumber, REPLACE(UPPER(CONVERT(varchar, [ASSIGNMENT_DATE], 106)), ' ', '/') AS 'Dispense', ASSIGNED_BY , IPLblShipExpiryDate, IPLblShipLotNo, KitRepled, SITEID, SUBJID,ROW_KEY FROM BIL_IP_RANGE WHERE (ROW_KEY = " + ROW_KEY + " AND SUBJID = '" + SUBJID + "') ORDER BY KitNumber";

            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlCommand cmd2 = new SqlCommand(sql2, con);
            var list = new Visits();
            var ipvisit = new List<IPStatus>();
            var visit = new List<VisitHistory>();
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var subjects = new IPStatus();
                    subjects.KitNumber = rdr["KitNumber"].ToString();
                    subjects.VISIT = rdr["VISIT"].ToString();
                    subjects.SITEID = rdr["SITEID"].ToString();
                    subjects.SUBJID = rdr["SUBJID"].ToString();
                    subjects.Dispense = rdr["Dispense"].ToString();
                    subjects.ASSIGNED_BY = rdr["ASSIGNED_BY"].ToString();
                    subjects.ROW_KEY = (int)rdr["ROW_KEY"];
                    subjects.IPLblShipExpiryDate = rdr["IPLblShipExpiryDate"].ToString();
                    subjects.IPLblShipLotNo = rdr["IPLblShipLotNo"].ToString();
                    subjects.KitRepled = rdr["KitRepled"].ToString();

                    ipvisit.Add(subjects);
                    TempData["SUBJID"] = SUBJID;
                    ViewBag.SPKEY = SPKEY;
                }
                rdr.Close();
                SqlDataReader rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    var subjects = new VisitHistory();
                    subjects.VISITKEY = (int)rdr2["VISITKEY"];
                    subjects.VISITID = rdr2["VISITID"].ToString();
                    subjects.SITEID = rdr2["SITEID"].ToString();
                    subjects.SUBJID = rdr2["SUBJID"].ToString();
                    subjects.ROW_KEY = (int)rdr2["ROW_KEY"];
                    subjects.ELIGIP = rdr2["ELIGIP"].ToString();


                    visit.Add(subjects);
                    TempData["SUBJID"] = SUBJID;
                    ViewBag.SPKEY = SPKEY;
                }
                rdr.Close();

            }
            con.Close();
            list.IPVisit = ipvisit;
            list.Visit = visit;
            return View(list);
        }

        [HttpPost]
        public IActionResult SubDispenseIP(Visits Request, int ROW_KEY, string SUBJID, string SITEID, string username, string password, int SPKEY, bool notdone )
        {
            string userid = HttpContext.Session.GetString("suserid");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

            string sql2 = "SELECT VISITKEY, VISITID, SITEID, SUBJID, ROW_KEY, ADDDATE, ADDUSER, ELIGIP FROM BIL_VISITS WHERE SUBJID = '" + SUBJID + "' AND ELIGIP = 'No' ORDER BY VISITKEY";
            string sql = "SELECT [VISIT], KitNumber, REPLACE(UPPER(CONVERT(varchar, [ASSIGNMENT_DATE], 106)), ' ', '/') AS 'Dispense', ASSIGNED_BY , IPLblShipExpiryDate, IPLblShipLotNo, KitRepled, SITEID, SUBJID,ROW_KEY FROM BIL_IP_RANGE WHERE (ROW_KEY = " + ROW_KEY + " AND SUBJID = '" + SUBJID + "') ORDER BY KitNumber";

            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlCommand cmd2 = new SqlCommand(sql2, con);
            var list = new Visits();
            var ipvisit = new List<IPStatus>();
            var visit = new List<VisitHistory>();
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var subjects = new IPStatus();
                    subjects.KitNumber = rdr["KitNumber"].ToString();
                    subjects.VISIT = rdr["VISIT"].ToString();
                    subjects.SITEID = rdr["SITEID"].ToString();
                    subjects.SUBJID = rdr["SUBJID"].ToString();
                    subjects.Dispense = rdr["Dispense"].ToString();
                    subjects.ASSIGNED_BY = rdr["ASSIGNED_BY"].ToString();
                    subjects.ROW_KEY = (int)rdr["ROW_KEY"];
                    subjects.IPLblShipExpiryDate = rdr["IPLblShipExpiryDate"].ToString();
                    subjects.IPLblShipLotNo = rdr["IPLblShipLotNo"].ToString();
                    subjects.KitRepled = rdr["KitRepled"].ToString();

                    ipvisit.Add(subjects);
                    TempData["SUBJID"] = SUBJID;
                    ViewBag.SPKEY = SPKEY;
                }
                rdr.Close();
                SqlDataReader rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    var subjects = new VisitHistory();
                    subjects.VISITKEY = (int)rdr2["VISITKEY"];
                    subjects.VISITID = rdr2["VISITID"].ToString();
                    subjects.SITEID = rdr2["SITEID"].ToString();
                    subjects.SUBJID = rdr2["SUBJID"].ToString();
                    subjects.ROW_KEY = (int)rdr2["ROW_KEY"];
                    subjects.ELIGIP = rdr2["ELIGIP"].ToString();


                    visit.Add(subjects);
                    TempData["SUBJID"] = SUBJID;
                    ViewBag.SPKEY = SPKEY;
                }
                rdr.Close();

            }
            con.Close();
            list.IPVisit = ipvisit;
            list.Visit = visit;
            if (checkidpwd(username, password) && (string.Equals(userid, username, StringComparison.OrdinalIgnoreCase) ))
            {
                //if(checkSubject(SITEID, SUBJID, Request.BRTHDTC, Request.SEX, SPKEY) == 0)
                //{
                //    TempData["ErrorMessage"] = "Can not find Subject.";
                //    return View();
                //}
                var rtnval = ChkSub(SPKEY, SITEID, SUBJID, Request.BRTHDTC, Request.SEX);
                if (rtnval != "OK")
                {
                    TempData["ErrorMessage"] = rtnval;
                    return View(list);
                }
                if (checkDupe(SPKEY, ROW_KEY, Request.VISITID, SUBJID) == "Exist")
                {
                    TempData["ErrorMessage"] = "This visit was marked as completed previously.";
                    return View(list);
                }
                string value = checkVist(SPKEY, ROW_KEY, SUBJID, Request.VISITID);
                if (value != "OK")
                {
                    TempData["ErrorMessage"] = "All previous visits must be completed. Skipped visits must be entered by checking not done.";
                    return View(list);
                }
                if(Request.ELIGIP == "Yes"){
                    string getArm = GetArm(SPKEY, ROW_KEY);
                    string getExpiry = GetExpiry(SPKEY, SITEID, getArm);
                    if(getExpiry == "NF")
                    {
                        TempData["ErrorMessage"] = "Kindly acknowledge the receipt of kits prior to dispensation.";
                        return View(list);
                    }
                    string rtnVal;
                    if (notdone)
                    {
                      rtnVal = ReqSubnotdone(SPKEY, ROW_KEY, SITEID, SUBJID, userid, Request.VISITID, Request.ELIGIP, Request.ReaYes, getExpiry, getArm, notdone);
                    }
                    else
                       rtnVal = ReqSub(SPKEY, ROW_KEY, SITEID, SUBJID, userid, Request.VISITID, Request.ELIGIP,  Request.ReaNo, getExpiry, getArm, Request.VISITDTC);
                    string[] arr = rtnVal.Split("|");
                    if (arr[0] == "OK")
                    {
                        string val = CheckAutoResupply(SPKEY, SITEID);
                        if ((val == "") && CheckStudyAutoRe(SPKEY) == "Enabled")
                        {
                            string result = ChkSiteInv(SPKEY, SITEID, getArm);
                            TempData["Message"] = arr[1]  ;
                            return RedirectToAction("SubsequentDispensationHome");
                        }
                        else
                        {
                            TempData["Message"] = arr[1];
                            return RedirectToAction("SubsequentDispensationHome");
                        }
                        

                    }
                    else
                    {
                        TempData["ErrorMessage"] = rtnVal;
                        return View(list);
                    }

                    //Check stopautoresupply
                    
                    //check Inventory


                }
                else
                {
                    string retVal;
                    if (notdone && Request.ELIGIP == "No")
                    {
                       retVal = ReceiveVisitNotdone(SPKEY, ROW_KEY, userid, Request.VISITID, SITEID, SUBJID, Request.ELIGIP, notdone);
                    }
                    else
                     retVal = ReceiveVisitNo(SPKEY, ROW_KEY, userid, Request.VISITID, SITEID, SUBJID, Request.ELIGIP, Request.ReaNo, Request.VISITDTC);
                    if(retVal == "")
                    {

                        TempData["Message"] = " Visit was marked as complete. But subject is not eligible to recieve IP";
                        string retVal2 = "";
                        string othEmail = "NF";
                        string retSite = "";

                        //retVal2 = genAct.GetNotify(amarexDbConnStr, SPKEY, "Randomized");
                        othEmail = GetProfVpe(SPKEY, "AddRandEmails");
                        retSite = GetEmailByGrp(SPKEY, SITEID, "S");
                        if (retVal2 == "")
                        {
                            retVal2 = "jacobk@amarexcro.com";
                        }
                        if (othEmail != "NF")
                        {
                            retVal2 += "; " + othEmail;
                        }
                        if (retSite != "")
                        {
                            retVal2 += "; " + retSite;
                        }
                        connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

                        string Selectsql = "SELECT * FROM BIL_SUBJ WHERE (SPKEY = " + SPKEY + ") AND (ROW_KEY = " + ROW_KEY + ")";
                        SqlConnection conn = new SqlConnection(connectionString);
                        SqlCommand Selectcmd = new SqlCommand(Selectsql, conn);
                        string msgBody = "Protocol: " + StudyName() + Environment.NewLine;
                        msgBody += Environment.NewLine + "This email is to notify you that subsequent dispensation for the following subject has been requested, details below. " + Environment.NewLine;
                        using (conn)
                        {
                            conn.Open();
                            SqlDataReader rdr = Selectcmd.ExecuteReader();
                            while (rdr.Read())
                            {
                                // string msgBody = "For Webview IRT Library System -  Request Subsequent IP Dispensation " + Environment.NewLine;
                                msgBody += "Site: " + rdr["SITEID"].ToString() + Environment.NewLine;
                                msgBody += "Subject ID: " + rdr["SUBJID"].ToString() + Environment.NewLine;
                                msgBody += "Year of Birth: " + rdr["BRTHDTC"].ToString() + Environment.NewLine;
                                msgBody += "Sex: " + rdr["SEX"].ToString() + Environment.NewLine;
                                msgBody += "Informed consent date: " + rdr["ICDTC"].ToString() + Environment.NewLine;
                                msgBody += "Eligibility for Randomization: " + rdr["ELIGRAND"].ToString() + Environment.NewLine;
                                //msgBody += "Randomization #: " + dt.Rows[0]["RANDNUM"].ToString() + Environment.NewLine;
                                msgBody += "Requested by: " + userid + Environment.NewLine;
                                if(Request.VISITDTC != null)
                                   msgBody += "Request date: " + Request.VISITDTC + Environment.NewLine;
                                msgBody += "Visit: " + Request.VISITID + Environment.NewLine;
                                if (Request.notdone)
                                    msgBody += "Visit was not done" + Environment.NewLine;
                                if(Request.ReaNo != null)
                                    msgBody += "Reason, if IP is not dispensed: " + Request.ReaNo + Environment.NewLine;
                            }

                        }
                        conn.Close();
                        string subject = StudyName() + " - WebView RTSM - Subject ID: " + SUBJID + " - Subsequent dispensation ";
                        SendEmail(retVal2, subject, msgBody);
                        return RedirectToAction("SubsequentDispensationHome");
                    }
                    else
                    {
                        TempData["ErrorMessage"] = retVal;
                        return View(list);
                    }
                }

                //return RedirectToAction("SubsequentDispensationHome");
            }
            else
                TempData["ErrorMessage"] = "Invalid username or password.";
                return View(list);
        }

        public bool checkidpwd(string username, string password)
        {

            var rtnVal = "";
            SecSSO chkSSO2 = new SecSSO();
            rtnVal = chkSSO2.ChkIDPWSSO(username, password, HttpContext.Session.GetString("sesuriSSIS"), HttpContext.Session.GetString("sesinstanceID"), HttpContext.Session.GetString("sesSecurityKey"), HttpContext.Session.GetString("sesAmarexDb"));

            if (rtnVal.Equals("7103"))
            {

                return true;
            }
            else
            {
                return false;
            }
        }

        //public int checkSubject(string SiteID, string SUBJID, string Brth, string sex, int spkey )
        //{
        //    connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
        //    int count = 0;
        //    string sql = "SELECT COUNT(*) FROM BIL_SUBJ WHERE (STATUS_INFO = 'Randomized') AND (SITEID = '" + SiteID + "') AND (SUBJID = '" + SUBJID + "') AND (SPKEY = " + spkey + ") AND (BRTHDTC = '" + Brth + "') AND (SEX = '" + sex + "')";
        //    SqlConnection con = new SqlConnection(connectionString);
        //    SqlCommand cmd = new SqlCommand(sql, con);
        //    using (con)
        //    {
        //        con.Open();
        //        count = (int)cmd.ExecuteScalar();


        //    }
        //    con.Close();
            
        //    return count;
        //}

        public string ChkSub(int spkey, string siteid, string subjid, string yob, string sex)
        {
            string sqlState = "";
            var retVal = "OK";
            string brthdtc;
            string gender;
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                sqlState = "SELECT BRTHDTC, SEX FROM BIL_SUBJ WHERE (SITEID = '" + siteid;
                sqlState += "') AND ([STATUS_INFO] = 'Randomized') AND (SUBJID = '" + subjid + "') AND (SPKEY = " + spkey + ")";
                SqlCommand selectCmd = new SqlCommand(sqlState, conn);
                using (SqlDataReader reader = selectCmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        brthdtc = reader["BRTHDTC"].ToString();
                        gender = reader["SEX"].ToString();
                        if ((brthdtc != yob) && (sex != gender))
                        {
                            retVal = "Year of birth and sex are incorrect for subject " + subjid + ".";
                            return retVal;
                        }
                        if (brthdtc != yob)
                        {
                            // Handle the case where the BRTHDTC does not match the request
                            retVal = "Year of birth is incorrect for subject " + subjid + ".";
                            return retVal;
                        }
                        if (sex != gender)
                        {
                            retVal = "Sex is incorrect for subject " + subjid + ".";
                            return retVal;
                        }

                    }
                    else
                    {
                        // Handle the case where the SUBJID is not found in BIL_SUBJ
                        retVal = "Subject ID " + subjid + " is not valid";
                        return retVal;
                    }
                }

            }
            return retVal;
        }

        public string checkDupe(int spkey, int rowkey, string visitid, string subjid)
        {
            string retVal = "OK";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
           
            string sql = "SELECT COUNT(*) FROM BIL_VISITS WHERE (SPKEY = " + spkey + ") AND (ROW_KEY = " + rowkey + ") AND (SUBJID = '" + subjid + "') AND (VISITID = '" + visitid + "')";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            using (con)
            {
                con.Open();
                
                int count = (int)cmd.ExecuteScalar();
                if (count >= 1)
                {
                    retVal = "Exist";
                }

            }
            con.Close();

            return retVal;
        }

        
        public string checkVist(int spkey, int rowkey, string subjid,  string selectedVisit)
        {
            string retVal = "OK";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            
            if (selectedVisit == "Visit 8")
            {
                string sql = "SELECT COUNT(*) FROM BIL_VISITS WHERE (SPKEY = " + spkey + ") AND (ROW_KEY = " + rowkey + ") AND (SUBJID = '" + subjid + "') AND (VISITID = 'Visit 6')";
                string sql1 = "SELECT COUNT(*) FROM BIL_VISITS WHERE (SPKEY = " + spkey + ") AND (ROW_KEY = " + rowkey + ") AND (SUBJID = '" + subjid + "') AND (VISITID = 'Visit 4')";
                int count = 0;
                int count1 = 0;
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        count = (int)cmd.ExecuteScalar();
                    }
                    using (SqlCommand cmd1 = new SqlCommand(sql1, con))
                    {
                        count1 = (int)cmd1.ExecuteScalar();
                    }
                    if (count >= 1 && count1 >= 1)
                    {
                        return retVal;
                    }
                    else if(count == 0 || count1 >= 1 )
                        return "Must have Visit 6 info";
                    else if (count >= 1 || count1 == 0)
                        return "Must have Visit 4 info";
                    else
                        return "Must have Visit 6 and Visit 4 info";
                }
            }

            else if (selectedVisit == "Visit 6")
            {
                string sql = "SELECT COUNT(*) FROM BIL_VISITS WHERE (SPKEY = " + spkey + ") AND (ROW_KEY = " + rowkey + ") AND (SUBJID = '" + subjid + "') AND (VISITID = 'Visit 4')";
                
                int count = 0;
                
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        count = (int)cmd.ExecuteScalar();
                    }

                    if (count >= 1)
                    {
                        return retVal;
                    }
                    else
                        return "Must have Visit 4 info";
                }
            }

            return retVal;
        }

        public string ReceiveVisitNo(int SPKEY, int ROW_KEY, string ADDUSER, string VISITID, string SITEID, string SUBJID, string ELIGIP, string reaNo, string VISITDTC)
        {
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string rtnVal = "";
            string sql = ""; 
            sql= "INSERT INTO BIL_VISITS (SPKEY, ROW_KEY, SITEID, SUBJID, VISITID, ELIGIP, ADDUSER, ReaNo, ADDDATE, VisitDate) VALUES (@SPKEY, @ROW_KEY, @SITEID, @SUBJID, @VISITID,  @ELIGIP, @ADDUSER, @reaNo, @ADDDATE, @VisitDate)";
            using(conn){
                conn.Open();
                SqlCommand insertCmd = new SqlCommand(sql, conn);
                insertCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                insertCmd.Parameters.AddWithValue("@ROW_KEY", ROW_KEY);
                insertCmd.Parameters.AddWithValue("@SITEID", SITEID);
                insertCmd.Parameters.AddWithValue("@SUBJID", SUBJID);
                insertCmd.Parameters.AddWithValue("@VISITID", VISITID);
                insertCmd.Parameters.AddWithValue("@ELIGIP", ELIGIP);
                insertCmd.Parameters.AddWithValue("@ADDUSER", ADDUSER);
                insertCmd.Parameters.AddWithValue("@ReaNo", (object)reaNo ?? DBNull.Value);
                insertCmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                insertCmd.Parameters.AddWithValue("VisitDate", VISITDTC);

                int rowsAfft = (int)insertCmd.ExecuteNonQuery();
                if (rowsAfft <= 0)
                {
                    rtnVal = "Error INSERT Rec Visit";
                }
            }
            conn.Close();
            return rtnVal;
        }

        public string ReceiveVisitNotdone(int SPKEY, int ROW_KEY, string ADDUSER, string VISITID, string SITEID, string SUBJID, string ELIGIP, bool notdone)
        {
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string rtnVal = "";
            string sql = "";
            sql = "INSERT INTO BIL_VISITS (SPKEY, ROW_KEY, SITEID, SUBJID, VISITID, ELIGIP, ADDUSER, notdone, ADDDATE, VisitDate) VALUES (@SPKEY, @ROW_KEY, @SITEID, @SUBJID, @VISITID,  @ELIGIP, @ADDUSER, @notdone, @ADDDATE, @VisitDate)";
            using (conn)
            {
                conn.Open();
                SqlCommand insertCmd = new SqlCommand(sql, conn);
                insertCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                insertCmd.Parameters.AddWithValue("@ROW_KEY", ROW_KEY);
                insertCmd.Parameters.AddWithValue("@SITEID", SITEID);
                insertCmd.Parameters.AddWithValue("@SUBJID", SUBJID);
                insertCmd.Parameters.AddWithValue("@VISITID", VISITID);
                insertCmd.Parameters.AddWithValue("@ELIGIP", ELIGIP);
                insertCmd.Parameters.AddWithValue("@ADDUSER", ADDUSER);
                insertCmd.Parameters.AddWithValue("@notdone", (object)notdone ?? DBNull.Value);
                insertCmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                insertCmd.Parameters.AddWithValue("@VisitDate", "Not Done");

                int rowsAfft = (int)insertCmd.ExecuteNonQuery();
                if (rowsAfft <= 0)
                {
                    rtnVal = "Error INSERT Rec Visit";
                }
            }
            conn.Close();
            return rtnVal;
        }

        public string GetArm(int SPKEY, int ROW_KEY)
        {
            SqlConnection con = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string rtnVal = "";
            string sql = "SELECT * FROM BIL_SUBJ WHERE (ROW_KEY = " + ROW_KEY + ") AND (SPKEY = " + SPKEY + ")";
            SqlCommand cmd = new SqlCommand(sql, con);
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    
                    rtnVal = rdr["ARMCD"].ToString();
                }
            }
            con.Close();
            return rtnVal;
        }

        public string GetExpiry(int SPKEY, string SITEID, string ARM)
        {
            SqlConnection con = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string rtnVal = "NF";
            //expiry date atleast 3 months from system's date
            //SELECT TOP 1 [IPLblShipExpiryDate],CAST([IPLblShipExpiryDate] AS DATE) AS chkExDT FROM BIL_IP_RANGE WHERE (SITEID = '01') AND ([TreatmentGroup] = 'ARM B') AND (RECVDBY IS NOT NULL) AND (ASSIGNED IS NULL) AND (KITSTAT = 'Acceptable')  AND (IPLblShipExpiryDate IS NOT NULL) AND ([IPLblShipExpiryDate] >= DATEADD(month, 3, GETDATE()))ORDER BY chkExDT;

            string sql = "SELECT TOP 1 [IPLblShipExpiryDate], CAST([IPLblShipExpiryDate] AS date) as chkExDT FROM BIL_IP_RANGE WHERE (SITEID = '" + SITEID + "') AND ([TreatmentGroup] = '" + ARM + "') AND (RECVDBY IS NOT NULL) AND (ASSIGNED IS NULL) AND (KITSTAT = 'Acceptable') AND (IPLblShipExpiryDate IS NOT NULL) ORDER BY chkExDT ";
            SqlCommand cmd = new SqlCommand(sql, con);
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {

                    rtnVal = rdr["IPLblShipExpiryDate"].ToString();
                    
                }
            }
            con.Close();
            return rtnVal;
        }

        public string ReqSub(int SPKEY, int ROW_KEY, string SITEID, string SUBJID, string uid, string VISITID, string eligIP, string reaNo, string expDt, string arm, string VISITDTC)
        {
            //Have to do notification Part
            string rtnVal = "";
            string kitNumber = AllocateKits(SPKEY, SITEID, SUBJID, uid, VISITID, ROW_KEY, arm, expDt, VISITDTC);
            if(kitNumber == "NF")
            {
                rtnVal = "No Kit at Site";
            }
            else
            {
                string kitInfo = "";
                kitInfo += "Kit Number " + kitNumber + " should be dispensed.";
                rtnVal = "OK|" + kitInfo;
                string Recvisit = RecVisit(SPKEY, ROW_KEY, uid,VISITID, kitNumber, eligIP, reaNo, SITEID, SUBJID, arm, VISITDTC);
                if(Recvisit == "")
                {
                    //Send Notification
                    string retVal2 = "";
                    string othEmail = "NF";
                    string retSite = "";
                    
                    //retVal2 = genAct.GetNotify(amarexDbConnStr, SPKEY, "Randomized");
                    othEmail = GetProfVpe(SPKEY, "AddRandEmails");
                    retSite = GetEmailByGrp( SPKEY, SITEID, "S");
                    if (retVal2 == "")
                    {
                        retVal2 = "jacobk@amarexcro.com";
                    }
                    if (othEmail != "NF")
                    {
                        retVal2 += "; " + othEmail;
                    }
                    if (retSite != "")
                    {
                        retVal2 += "; " + retSite;
                    }
                    connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                   
                    string sql = "SELECT * FROM BIL_SUBJ WHERE (SPKEY = " + SPKEY + ") AND (ROW_KEY = " + ROW_KEY + ")";
                    SqlConnection con = new SqlConnection(connectionString);
                    SqlCommand cmd = new SqlCommand(sql, con);
                    string msgBody = "Protocol: " + StudyName() + Environment.NewLine; 
                    msgBody += Environment.NewLine + "This email is to notify you that subsequent dispensation for the following subject has been requested, details below. " + Environment.NewLine; 
                    using (con)
                    {
                        con.Open();
                        SqlDataReader rdr = cmd.ExecuteReader();
                        while (rdr.Read())
                        {
                           // string msgBody = "For Webview IRT Library System -  Request Subsequent IP Dispensation " + Environment.NewLine;
                            msgBody += "Site: " + rdr["SITEID"].ToString() + Environment.NewLine;
                            msgBody += "Subject ID: " + rdr["SUBJID"].ToString() + Environment.NewLine;
                            msgBody += "Year of Birth: " + rdr["BRTHDTC"].ToString() + Environment.NewLine;
                            msgBody += "Sex: " + rdr["SEX"].ToString() + Environment.NewLine;
                            msgBody += "Informed consent date: " + rdr["ICDTC"].ToString() + Environment.NewLine;
                            msgBody += "Eligibility for Randomization: " + rdr["ELIGRAND"].ToString() + Environment.NewLine;
                            //msgBody += "Randomization #: " + dt.Rows[0]["RANDNUM"].ToString() + Environment.NewLine;
                            msgBody += "Requested by: " + uid + Environment.NewLine;
                            msgBody += "Request date: " + VISITDTC + Environment.NewLine;
                            msgBody += "Visit: " + VISITID + Environment.NewLine;
                            msgBody += "KitNumber: " + kitNumber + Environment.NewLine;
                        }

                    }
                    con.Close();
                    string subject = StudyName() + " - WebView RTSM - Subject ID: " + SUBJID + " - Subsequent dispensation ";
                    SendEmail(retVal2, subject, msgBody);
                }
                else
                {
                    rtnVal = Recvisit;
                }


            }

            return rtnVal;
        }
        public string ReqSubnotdone(int SPKEY, int ROW_KEY, string SITEID, string SUBJID, string uid, string VISITID, string eligIP, string reayes, string expDt, string arm, bool notdone)
        {
            //Have to do notification Part
            string rtnVal = "";
            string kitNumber = AllocateKitsnotdone(SPKEY, SITEID, SUBJID, uid, VISITID, ROW_KEY, arm, expDt, reayes, notdone);
            if (kitNumber == "NF")
            {
                rtnVal = "No Kit at Site";
            }
            else
            {
                string kitInfo = "";
                kitInfo += "Kit Number " + kitNumber + " should be dispensed.";
                rtnVal = "OK|" + kitInfo;
                string Recvisit = RecVisitNotdone(SPKEY, ROW_KEY, uid, VISITID, kitNumber, eligIP, reayes, SITEID, SUBJID, arm, notdone);
                if (Recvisit == "")
                {
                    //Send Notification
                    string retVal2 = "";
                    string othEmail = "NF";
                    string retSite = "";

                    //retVal2 = genAct.GetNotify(amarexDbConnStr, SPKEY, "Randomized");
                    othEmail = GetProfVpe(SPKEY, "AddRandEmails");
                    retSite = GetEmailByGrp(SPKEY, SITEID, "S");
                    if (retVal2 == "")
                    {
                        retVal2 = "jacobk@amarexcro.com";
                    }
                    if (othEmail != "NF")
                    {
                        retVal2 += "; " + othEmail;
                    }
                    if (retSite != "")
                    {
                        retVal2 += "; " + retSite;
                    }
                    connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

                    string sql = "SELECT * FROM BIL_SUBJ WHERE (SPKEY = " + SPKEY + ") AND (ROW_KEY = " + ROW_KEY + ")";
                    SqlConnection con = new SqlConnection(connectionString);
                    SqlCommand cmd = new SqlCommand(sql, con);
                    string msgBody = "Protocol: Webview RTSM - Test" + Environment.NewLine;
                    msgBody += Environment.NewLine + "This email is to notify you that subsequent dispensation for the following subject has been requested, details below. " + Environment.NewLine;
                    using (con)
                    {
                        con.Open();
                        SqlDataReader rdr = cmd.ExecuteReader();
                        while (rdr.Read())
                        {
                            // string msgBody = "For Webview IRT Library System -  Request Subsequent IP Dispensation " + Environment.NewLine;
                            msgBody += "Site: " + rdr["SITEID"].ToString() + Environment.NewLine;
                            msgBody += "Subject: " + rdr["SUBJID"].ToString() + Environment.NewLine;
                            msgBody += "Year of Birth: " + rdr["BRTHDTC"].ToString() + Environment.NewLine;
                            msgBody += "Sex: " + rdr["SEX"].ToString() + Environment.NewLine;
                            msgBody += "Informed consent date: " + rdr["ICDTC"].ToString() + Environment.NewLine;
                            msgBody += "Eligibility for Randomization: " + rdr["ELIGRAND"].ToString() + Environment.NewLine;
                            //msgBody += "Randomization #: " + dt.Rows[0]["RANDNUM"].ToString() + Environment.NewLine;
                            msgBody += "Requested by: " + uid + Environment.NewLine;
                            //msgBody += "Request date: " + DateTime.Now + Environment.NewLine;
                            msgBody += "Visit was not done " + Environment.NewLine;
                            msgBody += "Please provide the reason, if IP is dispensed " + reayes + Environment.NewLine;
                            msgBody += "Visit: " + VISITID + Environment.NewLine;
                            msgBody += "KitNumber: " + kitNumber + Environment.NewLine;
                            
                        }

                    }
                    con.Close();
                    string subject = StudyName() + " - WebView RTSM - Subject ID: " + SUBJID + " - Subsequent dispensation ";
                    SendEmail(retVal2, subject, msgBody);
                }
                else
                {
                    rtnVal = Recvisit;
                }


            }

            return rtnVal;
        }


        public string AllocateKits(int spkey, string siteid, string subjid, string uid, string visitid, int rowkey, string arm, string expDt, string VISITDTC)
        {
            string retVal = "NF";
            string sqlState = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DateTime dteNow = new DateTime();
            dteNow = DateTime.Now;
            sqlState = "UPDATE BIL_IP_RANGE SET ASSIGNED = 'Yes', ASSIGNMENT_DATE = '" + VISITDTC + "', ASSIGNED_BY = '" + uid + "', SUBJID = @SUBJID, [ROW_KEY] = @ROW_KEY, [VISIT] = @VISIT WHERE KITKEY IN ";
            //sqlState += "(SELECT TOP 1 KITKEY FROM BIL_IP_RANGE WHERE (SITEID = '" + siteid + "') AND ([TreatmentGroup] = '" + arm + "') AND (RECVDBY IS NOT NULL) AND (ASSIGNED IS NULL) ORDER BY [KitNumber]); ";
            sqlState += "(SELECT TOP 1 KITKEY FROM BIL_IP_RANGE WHERE (SITEID = '" + siteid + "') AND ([TreatmentGroup] = '" + arm + "') AND (RECVDBY IS NOT NULL) AND (ASSIGNED IS NULL) AND (KITSTAT = 'Acceptable') AND (IPLblShipExpiryDate = '" + expDt + "') ORDER BY [KitNumber]); ";
            string selectSql = "SELECT * FROM BIL_IP_RANGE WHERE (ASSIGNMENT_DATE = '" + VISITDTC + "') AND (ASSIGNED_BY = '" + uid + "') AND ([ROW_KEY] = " + rowkey + ") AND (SUBJID = '" + subjid + "') AND ([VISIT] = '" + visitid + "')";
          
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sqlState, conn);
            SqlCommand cmd2 = new SqlCommand(selectSql, conn);
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@SUBJID", subjid);
                cmd.Parameters.AddWithValue("@ROW_KEY", rowkey);
                cmd.Parameters.AddWithValue("@VISIT", visitid);
               
                cmd.ExecuteNonQuery();

                SqlDataReader rdr = cmd2.ExecuteReader();
                while (rdr.Read())
                {
                    
                    retVal = rdr["KitNumber"].ToString();

                }
                rdr.Close();

                //if (retVal != "NF")
                //{
                //    string sqlupdate = "Update BIL_IP_Range SET ASSIGNMENT_DATE = '" + VISITDTC + "' WHERE KITNUMBER = '" + retVal + "' ";
                //    SqlCommand cmd3 = new SqlCommand(sqlupdate, conn);
                //    cmd3.ExecuteNonQuery();
                //}
            }

            return retVal;

        }

        
        public string AllocateKitsnotdone(int spkey, string siteid, string subjid, string uid, string visitid, int rowkey, string arm, string expDt, string reayes, bool notdone)
        {
            string retVal = "NF";
            string sqlState = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DateTime dteNow = new DateTime();
            dteNow = DateTime.Now;
            sqlState = "UPDATE BIL_IP_RANGE SET ASSIGNED = 'Yes', ASSIGNMENT_DATE = '" + dteNow.ToString() + "', ASSIGNED_BY = '" + uid + "', SUBJID = @SUBJID, [ROW_KEY] = @ROW_KEY, [VISIT] = @VISIT, [ReaYes] = @ReaYes , [notdone] = @notdone WHERE KITKEY IN ";
            //sqlState += "(SELECT TOP 1 KITKEY FROM BIL_IP_RANGE WHERE (SITEID = '" + siteid + "') AND ([TreatmentGroup] = '" + arm + "') AND (RECVDBY IS NOT NULL) AND (ASSIGNED IS NULL) ORDER BY [KitNumber]); ";
            sqlState += "(SELECT TOP 1 KITKEY FROM BIL_IP_RANGE WHERE (SITEID = '" + siteid + "') AND ([TreatmentGroup] = '" + arm + "') AND (RECVDBY IS NOT NULL) AND (ASSIGNED IS NULL) AND (KITSTAT = 'Acceptable') AND (IPLblShipExpiryDate = '" + expDt + "') ORDER BY [KitNumber]); ";
            string selectSql = "SELECT * FROM BIL_IP_RANGE WHERE (ASSIGNMENT_DATE = '" + dteNow.ToString() + "') AND (ASSIGNED_BY = '" + uid + "') AND ([ROW_KEY] = " + rowkey + ")";
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sqlState, conn);
            SqlCommand cmd2 = new SqlCommand(selectSql, conn);
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@SUBJID", subjid);
                cmd.Parameters.AddWithValue("@ROW_KEY", rowkey);
                cmd.Parameters.AddWithValue("@VISIT", visitid);
                cmd.Parameters.AddWithValue("@ReaYes", reayes);
                cmd.Parameters.AddWithValue("@notdone", notdone);
                cmd.ExecuteNonQuery();

                SqlDataReader rdr = cmd2.ExecuteReader();
                while (rdr.Read())
                {

                    retVal = rdr["KitNumber"].ToString();

                }
                rdr.Close();

            }

            return retVal;

        }


        public string GetArmName(int SPKEY, int ROW_KEY, string SUBJID)
        {
            SqlConnection con = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string rtnVal = "";
            string sql = "SELECT * FROM BIL_SUBJ WHERE (ROW_KEY = " + ROW_KEY + ") AND (SPKEY = " + SPKEY + ")";
            SqlCommand cmd = new SqlCommand(sql, con);
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {

                    rtnVal = rdr["ARM"].ToString();
                }
            }
            con.Close();
            return rtnVal;
        }


        public string RecVisit(int SPKEY, int ROW_KEY, string ADDUSER, string VISITID, string KitNumber, string ELIGIP, string reaNo, string SITEID, string SUBJID, string ARM, string VISITDTC)
        {
            ARM = GetArmName(SPKEY, ROW_KEY, SUBJID);
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string rtnVal = "";
            string sql = "";
            sql = "INSERT INTO BIL_VISITS (SPKEY, ROW_KEY, SITEID, SUBJID, VISITID, ELIGIP, ADDUSER, ReaNo, KitNumber, ADDDATE, ARM, VisitDate) VALUES (@SPKEY, @ROW_KEY, @SITEID, @SUBJID, @VISITID,  @ELIGIP, @ADDUSER, @reaNo, @KitNumber, @ADDDATE, @ARM, @VisitDate)";
            using (conn)
            {
                conn.Open();
                SqlCommand insertCmd = new SqlCommand(sql, conn);
                insertCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                insertCmd.Parameters.AddWithValue("@ROW_KEY", ROW_KEY);
                insertCmd.Parameters.AddWithValue("@SITEID", SITEID);
                insertCmd.Parameters.AddWithValue("@SUBJID", SUBJID);
                insertCmd.Parameters.AddWithValue("@VISITID", VISITID);
                insertCmd.Parameters.AddWithValue("@ELIGIP", ELIGIP);
                insertCmd.Parameters.AddWithValue("@ADDUSER", ADDUSER);
                insertCmd.Parameters.AddWithValue("@KitNumber", KitNumber);
                insertCmd.Parameters.AddWithValue("@ReaNo", (object)reaNo ?? DBNull.Value);
                insertCmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                insertCmd.Parameters.AddWithValue("@ARM", ARM);
                insertCmd.Parameters.AddWithValue("@VisitDate", VISITDTC);

                int rowsAfft = (int)insertCmd.ExecuteNonQuery();
                if (rowsAfft <= 0)
                {
                    rtnVal = "Error Insert Rec Visit";
                }
            }
            conn.Close();
            return rtnVal;
        }

        public string RecVisitNotdone(int SPKEY, int ROW_KEY, string ADDUSER, string VISITID, string KitNumber, string ELIGIP, string reaYes, string SITEID, string SUBJID, string ARM, bool notdone)
        {
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            ARM = GetArmName(SPKEY, ROW_KEY, SUBJID);
            string rtnVal = "";
            string sql = "";
            sql = "INSERT INTO BIL_VISITS (SPKEY, ROW_KEY, SITEID, SUBJID, VISITID, ELIGIP, ADDUSER, ReaYes, KitNumber, ADDDATE, ARM, notdone, VisitDate) VALUES (@SPKEY, @ROW_KEY, @SITEID, @SUBJID, @VISITID,  @ELIGIP, @ADDUSER, @reaYes, @KitNumber, @ADDDATE, @ARM, @notdone, @VisitDate)";
            using (conn)
            {
                conn.Open();
                SqlCommand insertCmd = new SqlCommand(sql, conn);
                insertCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                insertCmd.Parameters.AddWithValue("@ROW_KEY", ROW_KEY);
                insertCmd.Parameters.AddWithValue("@SITEID", SITEID);
                insertCmd.Parameters.AddWithValue("@SUBJID", SUBJID);
                insertCmd.Parameters.AddWithValue("@VISITID", VISITID);
                insertCmd.Parameters.AddWithValue("@ELIGIP", ELIGIP);
                insertCmd.Parameters.AddWithValue("@ADDUSER", ADDUSER);
                insertCmd.Parameters.AddWithValue("@KitNumber", KitNumber);
                insertCmd.Parameters.AddWithValue("@reaYes", (object)reaYes ?? DBNull.Value);
                insertCmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                insertCmd.Parameters.AddWithValue("@ARM", ARM);
                insertCmd.Parameters.AddWithValue("@notdone", notdone);
                insertCmd.Parameters.AddWithValue("VisitDate", "Not Done");
                //insertCmd.Parameters.AddWithValue("@ReaNo", reaNo);

                int rowsAfft = (int)insertCmd.ExecuteNonQuery();
                if (rowsAfft <= 0)
                {
                    rtnVal = "Error Insert Rec Visit";
                }
            }
            conn.Close();
            return rtnVal;
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


        //IP Replacement 

        //public IActionResult IPReplacement(int ROW_KEY, string SUBJID, string SITEID, string SPKEY, string VISIT, string KitNumber)
        //{
        //    ViewBag.SITEID = SITEID;
        //    ViewBag.SUBJID = SUBJID;
        //    ViewBag.ROWKEY = ROW_KEY;
        //    TempData["SUBJID"] = SUBJID;
        //    TempData["ROW_KEY"] = ROW_KEY;
        //    return View();
        //}

        //[HttpPost]
        //public IActionResult IPReplacement(Visits Request, int ROW_KEY, string SUBJID, string SITEID, string username, string password, int SPKEY, string VISIT, string KitNumber)
        //{
        //    string userid = HttpContext.Session.GetString("suserid");
        //    TempData["SUBJID"] = SUBJID;
        //    TempData["ROW_KEY"] = ROW_KEY;
        //    if (checkidpwd(username, password))
        //    {
        //        if (checkSubject(SITEID, SUBJID, Request.BRTHDTC, Request.SEX, SPKEY) == 0)
        //        {
        //            TempData["ErrorMessage"] = "Cannot find Subject.";
        //            return View();
        //        }
        //        //string PreviousVisit = checkReplaceKit(SPKEY, ROW_KEY, SUBJID, SITEID, VISIT);
        //        //if(PreviousVisit == "Exist" && Re)
        //        string LastVisit = GetLastVisit(SPKEY, ROW_KEY);
        //        if (LastVisit != VISIT)
        //        {
        //            TempData["ErrorMessage"] = "Only the most recent visit kit can be replaced";
        //            return View();
        //        }
        //        string recentKitNumber = checkReplaceKit(SPKEY, ROW_KEY, SUBJID, SITEID, VISIT);
        //        if(recentKitNumber != KitNumber)
        //        {
        //            TempData["ErrorMessage"] = "IP Kit can't be replaced, only the most recently assigned IP kit can be replaced ";
        //            return View();
        //        }
        //        string getArm = GetArm(SPKEY, ROW_KEY);
        //           string getExpiry = GetExpiry(SPKEY, SITEID, getArm);
        //           if (getExpiry == "NF" || getExpiry == "")
        //            {
        //                TempData["ErrorMessage"] = "No IP for ARM and Expiry Date.";
        //                return View();
        //            }
        //            string rtnVal = ReqReplacement(SPKEY, ROW_KEY, SITEID, SUBJID, userid, Request.VISIT, KitNumber, Request.ReasonRep, getExpiry, getArm);
        //            string[] arr = rtnVal.Split("|");
        //        if (arr[0] == "OK")
        //        {
        //            string val = CheckAutoResupply(SPKEY, SITEID);
        //            if (val == "")
        //            {
        //                string result = ChkSiteInv(SPKEY, SITEID, getArm);
        //                TempData["Message"] = arr[1];
        //                return RedirectToAction("SubsequentDispensationHome");
        //            }
        //            else
        //            {
        //                TempData["Message"] = arr[1];
        //                return RedirectToAction("SubsequentDispensationHome");
        //            }


        //        }
        //        else
        //            {
        //                TempData["ErrorMessage"] = rtnVal;
        //                return View();
        //            }
        //    }
        //    else
        //        TempData["ErrorMessage"] = "Invalid username or password.";
        //    return View();
        //}

        //public string GetLastVisit(int SPKEY, int ROWKEY)
        //{
        //    SqlConnection con = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
        //    string rtnVal = "None";
        //    string sql = "SELECT TOP 1 VISITID FROM [BIL_VISITS] WHERE (SPKey = " + SPKEY + ") AND ([ROW_KEY] = " + ROWKEY + ") ORDER BY VISITID DESC";
        //    SqlCommand cmd = new SqlCommand(sql, con);
        //    using (con)
        //    {
        //        con.Open();
        //        SqlDataReader rdr = cmd.ExecuteReader();
        //        while (rdr.Read())
        //        {

        //            rtnVal = rdr["VISITID"].ToString();
        //        }
        //    }
        //    con.Close();
        //    return rtnVal;
        //}

        //public string ReqReplacement(int SPKEY, int ROW_KEY, string SITEID, string SUBJID, string userid, string VISITID, string KitNumber, string ReasonRep, string getExpiry, string getArm)
        //{
        //    string rtnVal = "";
        //    string kitNumber = AllocateReplacementKit(SPKEY, SITEID, SUBJID, userid, VISITID, ROW_KEY, getArm, getExpiry, KitNumber, ReasonRep);
        //    if (kitNumber == "NF" || kitNumber == "")
        //    {
        //        rtnVal = "No Kit at Site";
        //    }
        //    else
        //    {
        //        string kitInfo = "";
        //        kitInfo += "KitNumber " + kitNumber + " should be used.";
        //        rtnVal = "OK|" + kitInfo;
        //        //Send Notification
        //        string retVal2 = "";
        //        string othEmail = "NF";
        //        string retSite = "";

        //        //retVal2 = genAct.GetNotify(amarexDbConnStr, SPKEY, "Randomized");
        //        othEmail = GetProfVpe(SPKEY, "AddlRandEmails");
        //        retSite = GetEmailByGrp(SPKEY, SITEID, "S");
        //        if (retVal2 == "")
        //        {
        //            retVal2 = "jacobk@amarexcro.com";
        //        }
        //        if (othEmail != "NF")
        //        {
        //            retVal2 += "; " + othEmail;
        //        }
        //        if (retSite != "")
        //        {
        //            retVal2 += "; " + retSite;
        //        }
        //        connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

        //        string sql = "SELECT * FROM BIL_SUBJ WHERE (SPKEY = " + SPKEY + ") AND (ROW_KEY = " + ROW_KEY + ")";
        //        SqlConnection con = new SqlConnection(connectionString);
        //        SqlCommand cmd = new SqlCommand(sql, con);
        //        string msgBody = "Protocol: Webview RTSM - Test" + Environment.NewLine;
        //        msgBody += "For Webview RTSM -  Request Subsequent IP Dispensation " + Environment.NewLine; ;
        //        using (con)
        //        {
        //            con.Open();
        //            SqlDataReader rdr = cmd.ExecuteReader();
        //            while (rdr.Read())
        //            {
        //                // string msgBody = "For Webview IRT Library System -  Request Subsequent IP Dispensation " + Environment.NewLine;
        //                msgBody += "Site: " + rdr["SITEID"].ToString() + Environment.NewLine;
        //                msgBody += "Subject: " + rdr["SUBJID"].ToString() + Environment.NewLine;
        //                msgBody += "Year of Birth: " + rdr["BRTHDTC"].ToString() + Environment.NewLine;
        //                msgBody += "Sex: " + rdr["SEX"].ToString() + Environment.NewLine;
        //                msgBody += "Informed consent date: " + rdr["ICDTC"].ToString() + Environment.NewLine;
        //                msgBody += "Eligibility for Randomization: " + rdr["ELIGRAND"].ToString() + Environment.NewLine;
        //                //msgBody += "Randomization #: " + dt.Rows[0]["RANDNUM"].ToString() + Environment.NewLine;
        //                msgBody += "Visit: " + VISITID + Environment.NewLine;
        //                msgBody += "Reason: " + ReasonRep + Environment.NewLine;
        //            }

        //        }
        //        con.Close();
        //        string subject = "WebView RTSM IP Replacement for [Amarex][webview RTSM - Test] - Site " + SITEID + " - Subject " + SUBJID;
        //        SendEmail(retVal2, subject, msgBody);
                


        //    }
        //    return rtnVal;
        //}

        //public string AllocateReplacementKit(int SPKEY, string SITEID, string SUBJID, string userid, string VISITID, int ROW_KEY, string getArm, string getExpiry, string KitNumber, string ReasonRep)
        //{
        //    string retVal = "NF";
        //    string sqlState = "";
        //    connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
        //    DateTime dteNow = new DateTime();
        //    dteNow = DateTime.Now;
        //   // KitNumber = checkReplaceKit(SPKEY, ROW_KEY, SUBJID, SITEID, VISITID);
        //    sqlState = "UPDATE BIL_IP_RANGE SET ASSIGNED = 'Yes', ASSIGNMENT_DATE = '" + dteNow.ToString() + "', ASSIGNED_BY = '" + userid + "', SUBJID = '" + SUBJID + "', [ROW_KEY] = " + ROW_KEY + ", [VISIT] = '" + VISITID + "', [KITCOMM] = 'Replacement - Replaces: " + KitNumber + " - Reason: " + ReasonRep + "', KitRepled = '" + KitNumber + "' WHERE KITKEY IN ";
        //    sqlState += "(SELECT TOP 1 KITKEY FROM BIL_IP_RANGE WHERE (SITEID = '" + SITEID + "') AND (RECVDBY IS NOT NULL) AND (ASSIGNED IS NULL) AND ([TreatmentGroup] = '" + getArm + "') AND (IPLblShipExpiryDate = '" + getExpiry + "') ORDER BY [KitNumber]); ";
        //    string selectSql = "SELECT * FROM BIL_IP_RANGE WHERE (ASSIGNMENT_DATE = '" + dteNow.ToString() + "') AND (ASSIGNED_BY = '" + userid + "') AND ([ROW_KEY] = " + ROW_KEY + ")";
        //    SqlConnection conn = new SqlConnection(connectionString);
        //    SqlCommand cmd = new SqlCommand(sqlState, conn);
        //    SqlCommand cmd2 = new SqlCommand(selectSql, conn);
        //    using (conn)
        //    {
        //        conn.Open();
        //        cmd.ExecuteNonQuery();

        //        SqlDataReader rdr = cmd2.ExecuteReader();
        //        while (rdr.Read())
        //        {

        //            retVal = rdr["KitNumber"].ToString();

        //        }
        //        rdr.Close();
        //    }
        //    return retVal;
        //}
        //public string checkReplaceKit(int SPKEY, int ROWKEY, string SUBJID, string SITEID, string Visit)
        //{
        //    SqlConnection con = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
        //    string rtnVal = "None";
        //    string sql = "SELECT * FROM BIL_IP_RANGE WHERE (ROW_KEY = "+ROWKEY+") AND (SUBJID = '" + SUBJID+ "') AND (VISIT = '" + Visit + "') AND (SITEID = '" + SITEID + "') ORDER BY ASSIGNMENT_DATE DESC";
        //    SqlCommand cmd = new SqlCommand(sql, con);
        //    using (con)
        //    {
        //        con.Open();
        //        SqlDataReader rdr = cmd.ExecuteReader();
        //        while (rdr.Read())
        //        {

        //            rtnVal = rdr["KitNumber"].ToString();
        //            return rtnVal;
        //        }
        //    }
        //    con.Close();

        //    return rtnVal;

        //}

        public string GetProfVpe(int SPKEY, string type)
        {
            SqlConnection con = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string rtnVal = "";
            string sql = "SELECT * FROM Email_Notifications WHERE (SPKEY = " + SPKEY + ") AND (PIType = '" + type + "' AND Enable = 'True')";
            SqlCommand cmd = new SqlCommand(sql, con);
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rtnVal == "")
                    {
                        rtnVal = rdr["PIDet"].ToString();
                    }
                    else
                    {
                        rtnVal += ";" + rdr["PIDet"].ToString();
                    }
                }
            }
            con.Close();

            return rtnVal;
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


        public string ChkSiteInv(int spkey, string siteid, string arm)
        {
            string rtnVal = "";
            try
            {
                string sqlState;
                sqlState = "SELECT COUNT(*) AS CHKKITS FROM BIL_IP_RANGE WHERE (SITEID = '" + siteid + "') AND (TreatmentGroup = '" + arm + "') AND (ASSIGNED IS NULL) AND (KITSTAT = 'Acceptable') AND (RECVDBY IS NOT NULL)";
                SqlConnection con = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
                SqlCommand cmd = new SqlCommand(sqlState, con);
                con.Open();
                int cntArm = (int)cmd.ExecuteScalar();
                if (cntArm <= 8)
                {
                    string sqlState1 = "SELECT *, DATEDIFF(dd,BIL_IP_RANGE.SENDDATE,getdate()) AS SENDDAYS FROM BIL_IP_RANGE WHERE (SITEID = '" + siteid + "') AND (TreatmentGroup = '" + arm + "') AND (KITSTAT = 'Shipped')";
                    SqlCommand cmd2 = new SqlCommand(sqlState1, con);
                    SqlDataReader reader = cmd2.ExecuteReader();
                    if (reader.Read())
                    {
                        if (Convert.ToInt16(reader["SENDDAYS"]) > 4)
                        {
                            string msgBody = "Protocol: Webview RTSM - Test" + Environment.NewLine; ;
                            msgBody += "For Webview RTSM - Site inventory low, Kit in Shipped status for over 4 days " + Environment.NewLine;
                            msgBody += "Site: " + siteid + Environment.NewLine + "Kit Set: " + reader["KITSET"].ToString();
                            string retSupp = "";
                            //retSupp = GetEmailByGrp(spkey, "(All)", "D");
                            if (retSupp == "")
                            {
                                SendEmail("jacobk@amarexcro.com", "Webview RTSM - Kit Shipment - Site " + siteid + " - Low Inv with Shipped status - No Supp", msgBody);
                            }
                            else
                            {
                                SendEmail(retSupp + ";" + "jacobk@amarexcro.com", "Webview RTSM - Kit Shipment - Site " + siteid + " - Low Inv with Shipped status", msgBody);
                            }
                        }
                    }
                    else
                    {
                        reader.Close();
                        DateTime dteNow = new DateTime();
                        dteNow = DateTime.Now;
                        string sqlState2 = "DECLARE @KS int; SELECT @KS = MAX(KITSET) FROM BIL_IP_RANGE; IF @KS is null SELECT @KS = 1 ELSE SELECT @KS = @KS + 1; ";
                        sqlState2 += "UPDATE BIL_IP_RANGE SET SITEID = '" + siteid + "', KITSET = @KS, SENDBY = 'SYSTEM', SENDDATE = '" + dteNow.ToString() + "', KITTYPE = 'Auto Re-supply', KITSTAT = 'Shipped' WHERE KITKEY IN ";
                        sqlState2 += "(SELECT TOP 8 KITKEY FROM BIL_IP_RANGE WHERE (TreatmentGroup = '" + arm + "') AND ([ATDEPOT] is not NULL) AND (SITEID IS NULL) ORDER BY KitNumber); ";
                        SqlCommand cmd3 = new SqlCommand(sqlState2, con);
                        int rowsAfft = (int)cmd3.ExecuteNonQuery();
                        if (rowsAfft <= 0)
                        {
                            SendEmail("jacobk@amarexcro.com", "Webview RTSM - Auto Re-supply - Site " + siteid + " -  Unable to find kits", "Unable to find kits for Treatment Group: " + arm);
                        }
                        else
                        {
                            string kitsSel = "";
                            string sql = "SELECT * FROM BIL_IP_RANGE WHERE (SITEID = '" + siteid + "') AND (SENDBY = 'SYSTEM') AND (SENDDATE = '" + dteNow.ToString() + "') AND (KITTYPE = 'Auto Re-supply') AND (KITSTAT = 'Shipped')";
                            SqlCommand cmdSQL = new SqlCommand(sql, con);

                            SqlDataReader rdr = cmdSQL.ExecuteReader();
                            while (rdr.Read())
                            {
                                kitsSel += rdr["KitNumber"].ToString() + Environment.NewLine;
                            }
                            if (kitsSel == "")
                            {
                                rtnVal = "Error - Problem with Re-supply Shipment process, unable to find kits after selection.";
                                SendEmail("jacobk@amarexcro.com", "Webview RTSM - Auto Re-supply - Site " + siteid + " -  Shipment process, unable to find kits", "Shipment process, unable to find kits after selection for Treatment Group: " + arm);
                            }
                            else
                            {
                                sqlState = "Auto Re-supply Shipment" + Environment.NewLine + Environment.NewLine + kitsSel;
                                rtnVal = SentKitShipEmail(spkey, siteid, sqlState);
                            }
                        }
                    }
                }
                con.Close();
            }
            catch (Exception exp)
            {
                throw exp;
            }
            return rtnVal;
        }


        public string SentKitShipEmail(int spkey, string siteid, string kits)
        {
            string rtnVal = "OK";
            try
            {
                string toEmail = "";
                string retSite = "";
                string retSupp = "";
                string shipperEmail = "";
                //retSite = GetEmailByGrp(spkey, siteid, "S");
                //retSupp = GetEmailByGrp(spkey, "(All)", "D");
                shipperEmail =GetProfVpe(spkey, "PMIPLblRel");
                toEmail = retSite + ";" + retSupp;
                if (shipperEmail != "NF")
                {
                    if (toEmail == "")
                    {
                        toEmail = shipperEmail;
                    }
                    else
                    {
                        toEmail += ";" + shipperEmail;
                    }
                }
                if (toEmail == "")
                {
                    rtnVal = "No emails avaliable to notify of Shipment.";
                }
                else
                {
                    DateTime dtmDate = DateTime.Now;
                    string msgBody;
                    msgBody = "Webview RTSM" + Environment.NewLine + Environment.NewLine;
                    msgBody += "Request to Release - package and ship" + Environment.NewLine + Environment.NewLine;
                    msgBody += dtmDate.ToString("D") + Environment.NewLine + Environment.NewLine;
                    msgBody += GetShipTo(spkey, siteid);
                    msgBody += "KitNumber in shipment: " + Environment.NewLine + kits + Environment.NewLine + Environment.NewLine;
                    if (toEmail == "")
                    {
                        toEmail = "jacobk@amarexcro.com";
                    }

                    SendEmail(toEmail +";"+"sidran@amarecro.com", "Webview RTSM - Request to Release - Site " + siteid, msgBody);
                }
            }
            catch (Exception exp)
            {
                throw exp;
            }
            return rtnVal;
        }

        public string GetShipTo(int SPKEY, string SiteID)
        {
            string retVal = "No Ship To Info Found.";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            string sql = "SELECT * FROM [ShipToSite] WHERE ([SPKEY] = " + SPKEY + ") AND ([SITEID] = '" + SiteID + "')";
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    retVal = "Site Number: " + SiteID + "\r\n";
                    retVal += "Investigator Name: " + rdr["INVNAME"].ToString() + Environment.NewLine;
                    retVal += "Ship To Name: " + rdr["SITENAME"].ToString() + Environment.NewLine;
                    retVal += "Shipping Address 1: " + rdr["ADDR1"].ToString() + Environment.NewLine;
                    retVal += "Shipping Address 2: " + rdr["ADDR2"].ToString() + Environment.NewLine;
                    retVal += "Shipping City: " + rdr["CITY"].ToString() + Environment.NewLine;
                    retVal += "Shipping State: " + rdr["STATE"].ToString() + Environment.NewLine;
                    retVal += "Shipping Zip Code: " + rdr["ZIPCODE"].ToString() + Environment.NewLine;
                    retVal += "Shipping Country: " + rdr["COUNTRY"].ToString() + Environment.NewLine;
                    retVal += "Shipping Phone: " + rdr["PHONE"].ToString() + Environment.NewLine;
                    retVal += "Shipping Fax: " + rdr["FAX"].ToString() + Environment.NewLine;
                    retVal += "Shipping E-mail: " + rdr["EMAIL"].ToString() + Environment.NewLine;
                    retVal += "Special Instructions: " + rdr["SpecialInstructions"].ToString() + Environment.NewLine;

                }
                return retVal;
            }
        }

        public string CheckAutoResupply(int SPKEY, string SITEID)
        {
            string rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql2 = "SELECT PIDet From Stop_Auto_Supply WHERE  ([SPKEY] = " + SPKEY + ") AND ([PIDet] = '" + SITEID + "')";
            SqlConnection con = new SqlConnection(connectionString);
            using (con)
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(sql2, con);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    rtnVal = rdr["PIDet"].ToString();

                }

            }
            return rtnVal;
        }

        public string CheckStudyAutoRe(int SPKEY)
        {
            string rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql2 = "SELECT PIDet From Stop_Auto_Supply WHERE  ([SPKEY] = " + SPKEY + ") AND ([PIType] = 'AutoResupply')";
            SqlConnection con = new SqlConnection(connectionString);
            using (con)
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(sql2, con);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    rtnVal = rdr["PIDet"].ToString();

                }

            }
            return rtnVal;
        }


    }

}

