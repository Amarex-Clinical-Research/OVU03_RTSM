using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Webview_IRT.Models;

namespace RTSM_OLSingleArm.Controllers
{
    public class ReplacementController : Controller
    {
        public string connectionString;
        readonly IConfiguration _configuration;
        public ReplacementController(IConfiguration configuration)
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
        public IActionResult IPReplacementHome()
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

        public IActionResult SubjectVisits(int ROW_KEY, string SUBJID)
        {

            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

            string sql2 = "SELECT VISITKEY, VISITID, SITEID, SUBJID, ROW_KEY, ADDDATE, ADDUSER, ELIGIP FROM BIL_VISITS WHERE SUBJID = '" + SUBJID + "' AND ELIGIP = 'No' ORDER BY VISITKEY";
            string sql = "SELECT [VISIT], KitNumber, REPLACE(UPPER(CONVERT(varchar, [ASSIGNMENT_DATE], 106)), ' ', '/') AS 'Dispense', ASSIGNED_BY , IPLblShipExpiryDate, IPLblShipLotNo, KitRepled, SITEID, SUBJID,ROW_KEY FROM BIL_IP_RANGE WHERE (ROW_KEY = " + ROW_KEY + " AND SUBJID = '" + SUBJID + "') ORDER BY KitNumber";

            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlCommand cmd2 = new SqlCommand(sql2, con);
            var list = new Subjvisit();
            var ipvisit = new List<IPStatus>();
            var visit = new List<Visits>();
            var model = new List<IPStatus>();
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
                    var subjects = new Visits();
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
            TempData["SUBJID"] = SUBJID;
            list.IPVisit = ipvisit;
            list.Visit = visit;
            //  TempData.Keep("SUBJID");
            return View(list);

        }

        //IP Replacement 

        public IActionResult IPReplacement(int ROW_KEY, string SUBJID, string SITEID, int SPKEY, string VISIT, string KitNumber)
        {
            SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            ViewBag.SITEID = SITEID;
            ViewBag.SUBJID = SUBJID;
            ViewBag.ROWKEY = ROW_KEY;
            TempData["SUBJID"] = SUBJID;
            TempData["ROW_KEY"] = ROW_KEY;
            ViewBag.SPKEY = SPKEY;
            return View();
        }

        [HttpPost]
        public IActionResult IPReplacement(Visits Request, int ROW_KEY, string SUBJID, string SITEID, string username, string password, int SPKEY, string VISIT, string KitNumber)
        {
            string userid = HttpContext.Session.GetString("suserid");
            SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            TempData["SUBJID"] = SUBJID;
            TempData["ROW_KEY"] = ROW_KEY;
           
           if (checkidpwd(username, password) && (string.Equals(userid, username, StringComparison.OrdinalIgnoreCase) ))
            {
                var chkVal = checkSubject(SITEID, SUBJID, Request.BRTHDTC, Request.SEX, SPKEY);
                if (chkVal != "NV")
                {
                    TempData["ErrorMessage"] = chkVal;
                    return View();
                }
                //string PreviousVisit = checkReplaceKit(SPKEY, ROW_KEY, SUBJID, SITEID, VISIT);
                //if(PreviousVisit == "Exist" && Re)
                string LastVisit = GetLastVisit(SPKEY, ROW_KEY);
                if (LastVisit != VISIT)
                {
                    TempData["ErrorMessage"] = "Only the most recent visit kit can be replaced.";
                    return View();
                }
                string recentKitNumber = checkReplaceKit(SPKEY, ROW_KEY, SUBJID, SITEID, VISIT);
                if (recentKitNumber != KitNumber)
                {
                    TempData["ErrorMessage"] = "Kit can't be replaced, only the most recently assigned IP/Kit can be replaced. ";
                    return View();
                }
                string getArm = GetArm(SPKEY, ROW_KEY);
                string getExpiry = GetExpiry(SPKEY, SITEID, getArm);
                if (getExpiry == "NF" )
                {
                    TempData["ErrorMessage"] = "Kindly acknowledge the receipt of kits prior to dispensation.";
                    return View();
                }
                string rtnVal = ReqReplacement(SPKEY, ROW_KEY, SITEID, SUBJID, userid, Request.VISIT, KitNumber, Request.ReasonRep, getExpiry, getArm);
                string[] arr = rtnVal.Split("|");
                if (arr[0] == "OK")
                {
                    string val = CheckAutoResupply(SPKEY, SITEID);
                    if ((val == "") && CheckStudyAutoRe(SPKEY) == "Enabled")
                    {
                        string result = ChkSiteInv(SPKEY, SITEID, getArm);
                        TempData["Message"] = arr[1];
                        return RedirectToAction("IPReplacementHome");
                    }
                    else
                    {
                        TempData["Message"] = arr[1];
                        return RedirectToAction("IPReplacementHome");
                    }


                }
                else
                {
                    TempData["ErrorMessage"] = rtnVal;
                    return View();
                }
            }
            else
                TempData["ErrorMessage"] = "Invalid username or password.";
            return View();
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

        public string checkSubject(string SiteID, string SUBJID, string Brth, string sex, int spkey)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sqlState = "";
            var retVal = "NV";
            string brthdtc;
            string gender;
            string sql = "SELECT BRTHDTC, SEX FROM BIL_SUBJ WHERE (STATUS_INFO = 'Randomized') AND (SITEID = '" + SiteID + "') AND (SUBJID = '" + SUBJID + "') AND (SPKEY = " + spkey + ") ";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            using (con)
            {
                con.Open();
                using (SqlDataReader reader = cmd.ExecuteReader())
                {

                    if (reader.Read())
                    {
                        brthdtc = reader["BRTHDTC"].ToString();
                        gender = reader["SEX"].ToString();
                        if ((brthdtc != Brth) && (sex != gender))
                        {
                            retVal = "Year of birth and sex are incorrect for Subject " + SUBJID + ".";
                            return retVal;
                        }
                        if (brthdtc != Brth)
                        {
                            // Handle the case where the BRTHDTC does not match the request
                            retVal = "Year of birth is incorrect for Subject " + SUBJID + ".";
                            return retVal;
                        }
                        if (sex != gender)
                        {
                            retVal = "Sex is incorrect for subject " + SUBJID + ".";
                            return retVal;
                        }

                    }
                    else
                    {
                        // Handle the case where the SUBJID is not found in BIL_SUBJ
                        retVal = "Subject ID " + SUBJID + " is not valid";
                        return retVal;
                    }
                }


            }
            con.Close();

            return retVal;
        }


        public string GetLastVisit(int SPKEY, int ROWKEY)
        {
            SqlConnection con = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string rtnVal = "None";
            string sql = "SELECT TOP 1 VISITID FROM [BIL_VISITS] WHERE (SPKey = " + SPKEY + ") AND ([ROW_KEY] = " + ROWKEY + ") ORDER BY VISITID DESC";
            SqlCommand cmd = new SqlCommand(sql, con);
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {

                    rtnVal = rdr["VISITID"].ToString();
                }
            }
            con.Close();
            return rtnVal;
        }

        public string checkReplaceKit(int SPKEY, int ROWKEY, string SUBJID, string SITEID, string Visit)
        {
            SqlConnection con = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string rtnVal = "None";
            string sql = "SELECT * FROM BIL_IP_RANGE WHERE (ROW_KEY = " + ROWKEY + ") AND (SUBJID = '" + SUBJID + "') AND (VISIT = '" + Visit + "') AND (SITEID = '" + SITEID + "') ORDER BY ASSIGNMENT_DATE DESC";
            SqlCommand cmd = new SqlCommand(sql, con);
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {

                    rtnVal = rdr["KitNumber"].ToString();
                    return rtnVal;
                }
            }
            con.Close();

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
            //string sql = "SELECT TOP 1 [IPLblShipExpiryDate], CAST([IPLblShipExpiryDate] AS date) as chkExDT FROM BIL_IP_RANGE WHERE (SITEID = '" + SITEID + "') AND ([TreatmentGroup] = '" + ARM + "') AND (RECVDBY IS NOT NULL) AND (ASSIGNED IS NULL) ORDER BY chkExDT ";
            string sql = "SELECT TOP 1 [IPLblShipExpiryDate], CAST([IPLblShipExpiryDate] AS date) as chkExDT FROM BIL_IP_RANGE WHERE (SITEID = '" + SITEID + "') AND ([TreatmentGroup] = '" + ARM + "') AND (RECVDBY IS NOT NULL) AND (ASSIGNED IS NULL) AND (KITSTAT = 'Acceptable') AND IPLblShipExpiryDate IS NOT NULL ORDER BY chkExDT ";
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

        public string ReqReplacement(int SPKEY, int ROW_KEY, string SITEID, string SUBJID, string userid, string VISITID, string RepKitNumber, string ReasonRep, string getExpiry, string getArm)
        {
            string rtnVal = "";
            string kitNumber = AllocateReplacementKit(SPKEY, SITEID, SUBJID, userid, VISITID, ROW_KEY, getArm, getExpiry, RepKitNumber, ReasonRep);
            if (kitNumber == "NF" || kitNumber == "")
            {
                rtnVal = "No Kit at Site";
            }
            else
            {
                string kitInfo = "";
                kitInfo += "KitNumber " + kitNumber + " should be dispensed.";
                rtnVal = "OK|" + kitInfo;
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
                string msgBody = "Protocol: " + StudyName() + Environment.NewLine;
                msgBody += Environment.NewLine + "This email is to notify you that kit replacement for the following subject has been requested, details below." + Environment.NewLine; ;
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
                        msgBody += "Requested By: " + userid + Environment.NewLine;
                        msgBody += "Request Date: " + DateTime.Now.ToString("dd/MMM/yyyy") + Environment.NewLine;
                        msgBody += "Reason for replacement: " + ReasonRep + Environment.NewLine;
                        msgBody += "Visit: " + VISITID + Environment.NewLine;
                        msgBody += "KitNumber: " + kitNumber + Environment.NewLine;
                        msgBody += "Replacement for: " + RepKitNumber + Environment.NewLine;
                    }

                }
                con.Close();
                string subject = StudyName() + " - WebView RTSM - Subject ID: " + SUBJID + " - Kit replacement";
                SendEmail(retVal2, subject, msgBody);



            }
            return rtnVal;
        }

        public string AllocateReplacementKit(int SPKEY, string SITEID, string SUBJID, string userid, string VISITID, int ROW_KEY, string getArm, string getExpiry, string KitNumber, string ReasonRep)
        {
            string retVal = "NF";
            string sqlState = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DateTime dteNow = new DateTime();
            dteNow = DateTime.Now;
            // KitNumber = checkReplaceKit(SPKEY, ROW_KEY, SUBJID, SITEID, VISITID);
            sqlState = "UPDATE BIL_IP_RANGE SET ASSIGNED = 'Yes', ASSIGNMENT_DATE = '" + dteNow.ToString() + "', ASSIGNED_BY = '" + userid + "', SUBJID = @SUBJID, [ROW_KEY] = @ROW_KEY, [VISIT] = @VISIT, [KITCOMM] = @KITCOMM, KitRepled = @KitRepled WHERE KITKEY IN ";
            sqlState += "(SELECT TOP 1 KITKEY FROM BIL_IP_RANGE WHERE (SITEID = '" + SITEID + "') AND (RECVDBY IS NOT NULL) AND (ASSIGNED IS NULL) AND (KITSTAT = 'Acceptable') AND ([TreatmentGroup] = '" + getArm + "') AND (IPLblShipExpiryDate = '" + getExpiry + "') ORDER BY [KitNumber]); ";
            string selectSql = "SELECT * FROM BIL_IP_RANGE WHERE (ASSIGNMENT_DATE = '" + dteNow.ToString() + "') AND (ASSIGNED_BY = '" + userid + "') AND ([ROW_KEY] = " + ROW_KEY + ")";
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sqlState, conn);
            SqlCommand cmd2 = new SqlCommand(selectSql, conn);
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@SUBJID", SUBJID);
                cmd.Parameters.AddWithValue("@ROW_KEY", ROW_KEY);
                cmd.Parameters.AddWithValue("@VISIT", VISITID);
                cmd.Parameters.AddWithValue("@KITCOMM", "Replacement - Replaces: " + KitNumber + "- Reason: " + ReasonRep);
                cmd.Parameters.AddWithValue("@KitRepled", KitNumber);
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
                            retSupp = GetEmailByGrp(spkey, "(All)", "D");
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
                shipperEmail = GetProfVpe(spkey, "PMIPLblRel");
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

                    SendEmail(toEmail + ";" + "sidran@amarecro.com", "Webview RTSM - Request to Release - Site " + siteid, msgBody);
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
