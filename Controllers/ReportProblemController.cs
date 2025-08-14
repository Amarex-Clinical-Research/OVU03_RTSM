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
    public class ReportProblemController : Controller
    {
        public string connectionString;
        readonly IConfiguration _configuration;
        public ReportProblemController(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        
        
        public IActionResult ReportProblem()
        {
            ViewBag.siteid = HttpContext.Session.GetString("sesCenter");
            return View();
        }

        public IActionResult ReportProblemPost(Reportproblem Request, string SITEID)
        {
            ViewBag.siteid = HttpContext.Session.GetString("sesCenter");
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
           // string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            //string selectSql = "SELECT COUNT(*) FROM BIL_SUBJ WHERE SUBJID = @subjid AND (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND SPKEY = @SPKEY "; 
            string selectSql = "SELECT COUNT(*) FROM BIL_SUBJ WHERE SUBJID = @subjid AND SITEID = @SITEID  AND SPKEY = @SPKEY "; 
            string insertSql = "INSERT INTO BIL_PRBLM ( SPKEY, SITEID, SUBJID, PROBLEM_DESC, ADDUSER, ADDDATE, PROBLEM_STATUS) VALUES ( @SPKEY, @siteid, @subjid, @problem_desc, @ADDUSER, @ADDDATE, @PROBLEM_STATUS)";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand selectCmd = new SqlCommand(selectSql, con);
            SqlCommand insertCmd = new SqlCommand(insertSql, con);
            using (con)
            {

                con.Open();
                var rtnVal = "";
                
                
               
                if(Request.SUBJID != null && Request.SITEID != null)
                {
                    selectCmd.Parameters.AddWithValue("@subjid", Request.SUBJID);
                    selectCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                    selectCmd.Parameters.AddWithValue("@SITEID", SITEID);
                    int count = (int)selectCmd.ExecuteScalar();
                    if (count == 0)
                    {
                        // Handle invalid subjid
                        TempData["ErrorMessage"] = "Subject ID not found, please entered the valid Site ID and Subject ID.";
                        return RedirectToAction("ReportProblem");
                    }

                    insertCmd.Parameters.AddWithValue("@subjid", Request.SUBJID);
                    insertCmd.Parameters.AddWithValue("@siteid", SITEID);
                    insertCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                    insertCmd.Parameters.AddWithValue("@problem_desc", Request.PROBLEM_DESC);
                    insertCmd.Parameters.AddWithValue("@ADDUSER", userid);
                    DateTime currentDateTime = DateTime.Now;
                    insertCmd.Parameters.AddWithValue("@ADDDATE", currentDateTime);
                    insertCmd.Parameters.AddWithValue("@PROBLEM_STATUS", "Active");
                    insertCmd.ExecuteNonQuery();
                    string Message = "The request for Subject ID:" + Request.SUBJID + " has been successfully submitted.";
                    TempData["Message"] = Message;
                }
                else
                {
                    if(Request.NotApp && Request.SiteNA)
                    {
                        string sql = "INSERT INTO BIL_PRBLM ( SPKEY, SITEID, SUBJID, PROBLEM_DESC, ADDUSER, ADDDATE, PROBLEM_STATUS) VALUES ( @SPKEY, @siteid, @subjid, @problem_desc, @ADDUSER, @ADDDATE, @PROBLEM_STATUS)";
                        SqlCommand cmd = new SqlCommand(sql, con);
                        insertCmd.Parameters.AddWithValue("@subjid", "NA");
                        insertCmd.Parameters.AddWithValue("@siteid", "NA");
                        insertCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                        insertCmd.Parameters.AddWithValue("@problem_desc", Request.PROBLEM_DESC);
                        insertCmd.Parameters.AddWithValue("@ADDUSER", userid);
                        DateTime currentDateTime = DateTime.Now;
                        insertCmd.Parameters.AddWithValue("@ADDDATE", currentDateTime);
                        insertCmd.Parameters.AddWithValue("@PROBLEM_STATUS", "Active");
                        insertCmd.ExecuteNonQuery();
                        string Message = "Problem report has been successfully submitted.";
                        TempData["Message"] = Message;

                    }
                    else if (Request.NotApp)
                    {
                        rtnVal = CheckSite(Request.SITEID);
                        if(rtnVal == "")
                        {
                            TempData["ErrorMessage"] = "Site ID is incorrect";
                            return RedirectToAction("ReportProblem");
                        }
                        string sql = "INSERT INTO BIL_PRBLM ( SPKEY, SITEID, SUBJID, PROBLEM_DESC, ADDUSER, ADDDATE, PROBLEM_STATUS) VALUES ( @SPKEY, @siteid, @subjid, @problem_desc, @ADDUSER, @ADDDATE, @PROBLEM_STATUS)";
                        SqlCommand cmd = new SqlCommand(sql, con);
                        insertCmd.Parameters.AddWithValue("@subjid", "NA");
                        insertCmd.Parameters.AddWithValue("@siteid", Request.SITEID);
                        insertCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                        insertCmd.Parameters.AddWithValue("@problem_desc", Request.PROBLEM_DESC);
                        insertCmd.Parameters.AddWithValue("@ADDUSER", userid);
                        DateTime currentDateTime = DateTime.Now;
                        insertCmd.Parameters.AddWithValue("@ADDDATE", currentDateTime);
                        insertCmd.Parameters.AddWithValue("@PROBLEM_STATUS", "Active");
                        insertCmd.ExecuteNonQuery();
                        string Message = "Problem Report has been successfully submitted.";
                        TempData["Message"] = Message;

                    }
                    else if (Request.SiteNA)
                    {
                        rtnVal = CheckSUbject(Request.SUBJID);
                        if (rtnVal == "")
                        {
                            TempData["ErrorMessage"] = "Subject ID is incorrect";
                            return RedirectToAction("ReportProblem");
                        }
                        string sql = "INSERT INTO BIL_PRBLM ( SPKEY, SITEID, SUBJID, PROBLEM_DESC, ADDUSER, ADDDATE, PROBLEM_STATUS) VALUES ( @SPKEY, @siteid, @subjid, @problem_desc, @ADDUSER, @ADDDATE, @PROBLEM_STATUS)";
                        SqlCommand cmd = new SqlCommand(sql, con);
                        insertCmd.Parameters.AddWithValue("@subjid", Request.SUBJID);
                        insertCmd.Parameters.AddWithValue("@siteid", "NA");
                        insertCmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                        insertCmd.Parameters.AddWithValue("@problem_desc", Request.PROBLEM_DESC);
                        insertCmd.Parameters.AddWithValue("@ADDUSER", userid);
                        DateTime currentDateTime = DateTime.Now;
                        insertCmd.Parameters.AddWithValue("@ADDDATE", currentDateTime);
                        insertCmd.Parameters.AddWithValue("@PROBLEM_STATUS", "Active");
                        insertCmd.ExecuteNonQuery();
                        string Message = "Problem Report has been successfully submitted.";
                        TempData["Message"] = Message;

                    }
                    else
                    {
                        TempData["ErrorMessage"] = "Error occured while processing the request.";
                    }
                }
                
            }
            con.Close();
            string message;
            string subject;
           
            message = "Protocol: Webview RTSM - Test" + "\n";
            message += "We have received your report. Please allow up to 48 hours for a reply.";
            message += "Problem Report Details: " + "\n";
            if(Request.SITEID != null)
               message += "Site ID: " + Request.SITEID + "\n";
            if (Request.SUBJID != null)
                message += "Subject ID: " + Request.SUBJID + "\n";
            message += "Description: " + Request.PROBLEM_DESC + "\n";
            subject = "Webview RTSM - Problem Report for [Amarex][WebView RTSM -Test] Subject " + Request.SUBJID;
            SendEmail(GetUserEmail(), subject, message);
           

            string emailadd = GetNotify(SPKEY, "Problem Reported") + ";" + "sidran@amarexcro.com  ;";
            message = "Protocol: Webview RTSM - Test" + "\n";
            message += "A problem report has been submitted from " + GetUserEmail() + " ." + "\n" ;
            if (Request.SITEID != null)
                message += "Site ID: " + Request.SITEID + "\n";
            if (Request.SUBJID != null)
                message += "Subject ID: " + Request.SUBJID + "\n";
            message += "Description: " + Request.PROBLEM_DESC + "\n";


            subject = "Webview RTSM - Problem Report for [Webview][WebView RTSM -Test]";
            SendEmail(emailadd, subject, message);

            return RedirectToAction("ReportProblem");

        }


        public string CheckSite(string SITEID)
        {
            string rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            string sql = "SELECT COUNT(*) FROM ShipToSite WHERE SITEID = @SITEID AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + "";
            SqlCommand cmd = new SqlCommand(sql, conn);
            int count = 0;
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@SITEID", SITEID);
                count = (int)cmd.ExecuteScalar();
                   
                conn.Close();
            }
            if (count >= 1)
                rtnVal = "Site Found";
            return rtnVal;
        }

        public string CheckSUbject(string SUBJID)
        {
            string rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            string sql = "SELECT COUNT(*) FROM BIL_SUBJ WHERE SUBJID = @SUBJID AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + "";
            SqlCommand cmd = new SqlCommand(sql, conn);
            int count = 0;
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@SUBJID", SUBJID);
                count = (int)cmd.ExecuteScalar();

                conn.Close();
            }
            if (count >= 1)
                rtnVal = "SUbject Found";
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


    }
}
