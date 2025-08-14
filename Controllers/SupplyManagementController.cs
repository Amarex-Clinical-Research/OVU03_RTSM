using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
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
    public class SupplyManagementController : Controller
    {
        private readonly ILogger<SupplyManagementController> _logger;
        public string connectionString;
        readonly IConfiguration _configuration;

        public SupplyManagementController(ILogger<SupplyManagementController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        // Receive IP Sub Module
       

        
        private DataTable CreateDataTable(string cmdText)
        {
            System.Data.DataTable dt = new DataTable();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            System.Data.SqlClient.SqlDataAdapter da = new SqlDataAdapter(cmdText, connectionString);
            da.Fill(dt);
            return dt;
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

        public string GetUserName(string userid)
        {
            string name = "";
            string email = "";
            string sql = "SELECT User_Email, LName, FName FROM zSecurityID WHERE (UserID = '" + userid + " ')";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("AmarexDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    email = rdr["User_Email"].ToString();
                    name = rdr["FName"].ToString();
                    name += " " + rdr["LName"].ToString();

                }
            }
            return name;
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

        


        public string ChkSiteInv(int spkey, string siteid, string arm)
        {
            string rtnVal = "";
            try
            {
                string sqlState;
                sqlState = "SELECT COUNT(*) AS CHKKITS FROM BIL_IP_RANGE WHERE (SITEID = '" + siteid + "') AND (TreatmentGroup = '" + arm + "') AND (KITSTAT = 'Acceptable') AND (ASSIGNED IS NULL) AND (RECVDBY IS NOT NULL)";
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
                            string msgBody;
                            msgBody = "For Webview RTSM - Site inventory low, Kit in Shipped status for over 4 days " + Environment.NewLine;
                            msgBody += "Site: " + siteid + Environment.NewLine + "Kit Set: " + reader["KITSET"].ToString();
                            string retSupp = "";
                            //retSupp = GetEmailByGrp(spkey, "(All)", "D");
                            if (retSupp == "")
                            {
                                SendEmail("sidran@amarexcro.com", "Webview RTSM - Kit Shipment - Site " + siteid + " - Low Inv with Shipped status - No Supp", msgBody);
                            }
                            else
                            {
                                SendEmail(retSupp + ";" + "sidran@amarexcro.com", "Webview RTSM - Kit Shipment - Site " + siteid + " - Low Inv with Shipped status", msgBody);
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
                            SendEmail("sidran@amarexcro.com", "Webview RTSM - Auto Re-supply - Site " + siteid + " -  Unable to find kits", "Unable to find kits for Treatment Group: " + arm);
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
                                SendEmail("sidran@amarexcro.com", "Webview RTSM - Auto Re-supply - Site " + siteid + " -  Shipment process, unable to find kits", "Shipment process, unable to find kits after selection for Treatment Group: " + arm);
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
                        toEmail = "sidran@amarexcro.com";
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
        //Receive IP for Shipment level
        public IActionResult ReceiveIPTabs()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            String sql = "SELECT [SITEID], [KITSET], [KITSTAT], REPLACE(UPPER(CONVERT(varchar, [SENDDATE], 106)), ' ', '/') AS 'SENDDATEstr'  FROM [BIL_IP_RANGE] WHERE (([SITEID] = '" + HttpContext.Session.GetString("sesCenter") + "'OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND ([SENDBY] IS NOT NULL) AND ([KITSTAT] = 'Shipped') AND ([ATDEPOT] is not NULL)) GROUP BY [SITEID],[KITSET],[KITSTAT], [SENDDATE]";
            String sql2 = "SELECT [SITEID], [KITSET] FROM [BIL_IP_RANGE] WHERE (([SITEID] = '" + HttpContext.Session.GetString("sesCenter") + "'OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND ([SENDBY] IS NOT NULL) AND([KITSTAT] = 'Acceptable' OR [KITSTAT] = 'Unacceptable')) GROUP BY [SITEID],[KITSET]";
            String sql3 = "SELECT [SITEID], [KitNumber], [KITSTAT], REPLACE(UPPER(CONVERT(varchar, [SENDDATE], 106)), ' ', '/') AS 'SEND DATE', REPLACE(UPPER(CONVERT(varchar, [RECVDDATE], 106)), ' ', '/') AS 'RECVD DATE', [ASSIGNED], [SUBJID], [KITKEY], [RECVDBY], [SENDBY], [TreatmentGroup] FROM [BIL_IP_RANGE] WHERE ([SITEID] = '" + HttpContext.Session.GetString("sesCenter") + "'OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND ([SENDBY] IS NOT NULL) AND ([KITSTAT] = 'Unacceptable') ORDER BY [SITEID], SUBJID, [KitNumber], [KITKEY]";
            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlCommand cmd2 = new SqlCommand(sql2, con);
            SqlCommand cmd3 = new SqlCommand(sql3, con);

            var list = new IPList();
            var ipstat = new List<IPStatus>();
            var shipip = new List<ShipIP>();
            var damage = new List<DamageKit>();

            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var temp = new IPStatus();
                    temp.KITSET = (short)rdr["KITSET"];
                    temp.SITEID = rdr["SITEID"].ToString();
                    temp.KITSTAT = rdr["KITSTAT"].ToString();
                    temp.SENDDATEstr = rdr["SENDDATEstr"].ToString();
                    //temp.SENDBY = rdr["SENDBY"].ToString();

                    ipstat.Add(temp);
                }
                rdr.Close();
                SqlDataReader rdr3 = cmd2.ExecuteReader();
                while (rdr3.Read())
                {
                    var temp = new ShipIP();
                    temp.KITSET = (short)rdr3["KITSET"];
                    temp.SITEID = rdr3["SITEID"].ToString();
                    //temp.KITSTAT = rdr3["KITSTAT"].ToString();
                    shipip.Add(temp);
                }
                rdr3.Close();
                SqlDataReader rdr4 = cmd3.ExecuteReader();
                while (rdr4.Read())
                {
                    var temp = new DamageKit();
                    temp.KITKEY = (int)rdr4["KITKEY"];
                    temp.SITEID = rdr4["SITEID"].ToString();
                    temp.KitNumber = rdr4["KitNumber"].ToString();
                    temp.KITSTAT = rdr4["KITSTAT"].ToString();
                    temp.TreatmentGroup = rdr4["TreatmentGroup"].ToString();
                    temp.SENDDATEstr = rdr4["SEND DATE"].ToString();
                    temp.RECVDDATEstr = rdr4["RECVD DATE"].ToString();
                    temp.RECVDBY = rdr4["RECVDBY"].ToString();
                    temp.SENDBY = rdr4["SENDBY"].ToString();
                    damage.Add(temp);
                }
                rdr3.Close();

            }
            con.Close();
            list.IPStatus = ipstat;
            list.ShipIP = shipip;
            list.DamageKit = damage;
            return View(list);
        }

        public IActionResult IPReceiptReporting(int KITSET, string SITEID, string SENDDATE)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            ViewBag.SiteID = SITEID;
            ViewBag.KITSET = KITSET;
            ViewBag.SENDDATE = SENDDATE;

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            string sql = "SELECT KITKEY, KitNumber, SITEID, KITSTAT, REPLACE(UPPER(CONVERT(varchar, [SENDDATE], 106)), ' ', '/') AS 'SENDDATEstr' FROM BIL_IP_RANGE WHERE KITSET = @KITKEY AND KITSTAT = 'Shipped'";
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
                    selectCmd.Parameters.AddWithValue("@KITKEY", KITSET);

                    SqlDataReader rdr = selectCmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var temp = new ShipIP();
                        temp.KITKEY = (int)rdr["KITKEY"];
                        temp.KitNumber = rdr["KitNumber"].ToString();
                        temp.SITEID = rdr["SITEID"].ToString();
                        temp.KITSTAT = rdr["KITSTAT"].ToString();
                        temp.SENDDATEstr = rdr["SENDDATEstr"].ToString();

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
        public IActionResult IPReceiptReporting(int KITSET, string SITEID, string SENDDATE, string SelectedKitKeys, IPReporting model, string RECVDDATE, string isExcrusion, string SPAPDTC, string Isapproved, string username, string password, string IsPhysicalDamage)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            ViewBag.SiteID = SITEID;
            ViewBag.KITSET = KITSET;
            ViewBag.SENDDATE = SENDDATE;
            string sql;
            string rtnVal = "";
            if (IsValidUser(username, password) && (string.Equals(userid, username, StringComparison.OrdinalIgnoreCase) || userid == "sidran" || userid == "test1"))
            {
                //If there is no temperature Excursion
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
                if(isExcrusion == "No")
                {
                    if(IsPhysicalDamage == "No")
                      rtnVal = IPReceive2(SITEID, KITSET, "Acceptable", RECVDDATE, isExcrusion, SPAPDTC, Isapproved, username, IsPhysicalDamage);
                    else 
                    {
                        if (SelectedKitKeys != null)
                        {
                            string[] kitKeysArray = SelectedKitKeys.Split(',');

                            if (CountKitSet(KITSET) == kitKeysArray.Length)
                            {
                                rtnVal = IPReceive2(SITEID, KITSET, "Acceptable", RECVDDATE, isExcrusion, SPAPDTC, Isapproved, username, IsPhysicalDamage);
                            }
                            else
                            {
                                rtnVal = UpdateRows(SITEID, SelectedKitKeys, KITSET, RECVDDATE, isExcrusion, SPAPDTC, Isapproved, username, CountKitSet(KITSET), IsPhysicalDamage);
                            }
                        }
                        else
                        {
                            rtnVal = IPProcessed(SITEID, KITSET, "Unacceptable", RECVDDATE, isExcrusion, SPAPDTC, Isapproved, username, IsPhysicalDamage);
                        }
                    }
                }
                else if(isExcrusion == "Yes" && Isapproved == "No")
                {
                    rtnVal = IPReceive2(SITEID, KITSET, "Unacceptable", RECVDDATE, isExcrusion, SPAPDTC, Isapproved, username, IsPhysicalDamage);
                }
                else if(isExcrusion == "Yes" && Isapproved == "Yes" && IsPhysicalDamage == "No")
                {
                    rtnVal = IPReceive2(SITEID, KITSET, "Acceptable", RECVDDATE, isExcrusion, SPAPDTC, Isapproved, username, IsPhysicalDamage);
                }
                else if(isExcrusion == "Yes" && Isapproved == "Yes" && IsPhysicalDamage == "Yes")
                {
                    if (SelectedKitKeys != null)
                    {
                        string[] kitKeysArray = SelectedKitKeys.Split(',');

                        if (CountKitSet(KITSET) == kitKeysArray.Length)
                        {
                            rtnVal = IPReceive2(SITEID, KITSET, "Acceptable", RECVDDATE, isExcrusion, SPAPDTC, Isapproved, username, IsPhysicalDamage);
                        }
                        else
                        {
                            rtnVal = UpdateRows(SITEID, SelectedKitKeys, KITSET, RECVDDATE, isExcrusion, SPAPDTC, Isapproved, username, CountKitSet(KITSET), IsPhysicalDamage);
                        }
                    }
                    else
                    {
                        rtnVal = IPProcessed(SITEID, KITSET, "Unacceptable", RECVDDATE, isExcrusion, SPAPDTC, Isapproved, username, IsPhysicalDamage);
                    }

                }
               
             if(rtnVal == "Error processing Kit at Site contact your PM/CRA.")
                {
                    TempData["ErrorMessage"] = "Error processing Kit at Site contact your PM/CRA.";
                    return View();
                }
             else if(rtnVal == "Kits has been processed" || rtnVal == "Unacceptable")
                {
                    if (rtnVal == "Unacceptable")
                        TempData["Message"] = "Shipment Key " + KITSET + " has been received in unacceptable condition. Please quarantine. ";
                    else
                        TempData["Message"] = "Part of the kit shipment has been received in acceptable condition. ";
                    
                }
               else if (rtnVal == "Acceptable")
                {
                    TempData["Message"] = "Shipment Key " + KITSET + " has been received in acceptable condition. ";

                }
                else if(rtnVal == "Error: Can not find a valid inventroy Type")
                {
                    TempData["ErrorMessage"] = rtnVal;
                }
                else
                {
                    TempData["ErrorMessage"] = "Error occured while processing the receipt reporting";
                }
                LogActivity(KITSET);
                //return RedirectToAction("ReceiveIPTabs");
                string val = CheckAutoResupply(SPKEY, SITEID);
                if ((val == "") && CheckStudyAutoRe(SPKEY) == "Enabled")
                {
                    string result = ChkSiteInv2(SPKEY, SITEID);

                }
                return RedirectToAction("UploadFiles", new { KITSET = KITSET, SITEID = SITEID });


            }
            else
            {
                // Invalid username or password
                string errorMessage = "Invalid username or password.";
                //ViewBag.ErrorMessage = errorMessage;
                TempData["ErrorMessage"] = errorMessage;
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
                sql = "SELECT KITKEY, KitNumber, SITEID, KITSTAT FROM BIL_IP_RANGE WHERE KITSET = @KITKEY AND KITSTAT = 'Shipped'";
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
                        selectCmd.Parameters.AddWithValue("@KITKEY", KITSET);

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
                //return View();
            }

        }

        public Byte[] FileBytes(IFormFile File)
        {
            Byte[] bytes = null;
            using (MemoryStream ms = new MemoryStream())
            {
                File.OpenReadStream().CopyTo(ms);
                bytes = ms.ToArray();
            }
            return bytes;
        }


        public string IPReceive2(string SITEID, int KITSET, string proctype, string RECVDDATE, string isExcrusion, string SPAPDTC, string Isapproved, string username, string IsPhysicalDamage)
        {
            string rtn = "";
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            //string SITEID = HttpContext.Session.GetString("sesCenter");
            string sql = "";
            //string recieve = "acceptable";
            if (isExcrusion == "Yes" && Isapproved == "Yes")
            {
                sql = "UPDATE BIL_IP_RANGE SET  RECVDBY = @RECVDBY, RECVDDATE = @RECVDDATE, KITSTAT = '" + proctype + "', isExcrusion = @isExcrusion, SPAPDTC = @SPAPDTC, IsPhysicalDamage = @IsPhysicalDamage  WHERE KITSET = @KITSET AND KITSTAT = 'Shipped'";
            }
            else if (proctype == "Acceptable")
            { 
                sql = "UPDATE BIL_IP_RANGE SET RECVDBY = @RECVDBY, RECVDDATE = @RECVDDATE, KITSTAT = '" + proctype + "', isExcrusion = @isExcrusion, IsPhysicalDamage = @IsPhysicalDamage  WHERE KITSET = @KITSET AND KITSTAT = 'Shipped'";
            }
            else if(proctype == "Unacceptable")
            {
                sql = "UPDATE BIL_IP_RANGE SET  RECVDBY = @RECVDBY, RECVDDATE = @RECVDDATE, KITSTAT = '" + proctype + "', isExcrusion = @isExcrusion, IsPhysicalDamage = @IsPhysicalDamage WHERE KITSET = @KITSET AND KITSTAT = 'Shipped' ";
                //recieve = "Unacceptable";
            }
            else
            {
                rtn = "Error: Can not find a valid inventroy Type";
                return rtn;
            }
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            int rowsAfft;
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@KITSET", KITSET);
                cmd.Parameters.AddWithValue("@RECVDBY", userid);
                cmd.Parameters.AddWithValue("@isExcrusion", isExcrusion);
                cmd.Parameters.AddWithValue("@RECVDDATE", RECVDDATE);
                if ((SPAPDTC != null) && (Isapproved == "Yes" && isExcrusion == "Yes"))
                {
                    cmd.Parameters.AddWithValue("@SPAPDTC", (object)SPAPDTC ?? DBNull.Value);
                }
               
                  cmd.Parameters.AddWithValue("@IsPhysicalDamage", (object)IsPhysicalDamage ?? DBNull.Value);
                rowsAfft = (int)cmd.ExecuteNonQuery();

            }
            if (rowsAfft == 0)
            {
                rtn = "Error processing Kit at Site contact your PM/CRA.";
            }
            else
            {
                if (proctype == "Acceptable")
                {
                    string UserName = GetUserName(HttpContext.Session.GetString("suserid"));
                    string subject = "Webview RTSM - Shipment Processed at Site " + SITEID + "";
                    string message = "Protocol: Webview RTSM - Test" + "\n" + "Shipment " + KITSET + " has been received in acceptable condition  by " + UserName + " at Site " + SITEID + "." + "\n" ;
                    message += "Received Date: " + RECVDDATE + "\n";
                    message += "Temperature Excursion: " + isExcrusion + "\n";
                    if (Isapproved != null)
                        message += "Sponsor approved the use of quarantined IP after temperature excursion?: " + Isapproved + "\n";
                    if (SPAPDTC != null)
                        message += "Date Of Sponsor Approval: " + SPAPDTC + "\n";
                    if (IsPhysicalDamage != null)
                        message += "Was there any Physical Damage: " + IsPhysicalDamage + "\n";
                    string email = GetEmail(HttpContext.Session.GetString("suserid"));
                    string retSupp = "";
                    string otheremail = "";
                    retSupp = GetEmailByGrp(SPKEY, "(All)", "D");
                    otheremail = GetEmailByGrp(SPKEY, SITEID, "S");
                    email = email + ";" + retSupp + ";" + otheremail;
                    SendEmail(email, subject, message);
                    rtn = "Acceptable";
                }
                if (proctype == "Unacceptable")
                {
                    string UserName = GetUserName(HttpContext.Session.GetString("suserid"));
                    string subject = "Webview RTSM - Shipment Processed for [Amarex][Webview RTSM - Test] at Site " + SITEID + "";
                    string message = "Protocol: Webview RTSM - Test " + "\n" + "Shipment ID: " + KITSET + " has been received in Unacceptable condition by " + UserName + " at Site " + SITEID + "." + "\n";
                    message += "Received Date: " + RECVDDATE + "\n";
                    message += "Temperature Excursion: " + isExcrusion + "\n";
                    if (Isapproved != null)
                        message += "Sponsor approved the use of quarantined IP after temperature excursion?: " + Isapproved + "\n";
                    if (SPAPDTC != null)
                        message += "Date Of Sponsor Approval: " + SPAPDTC + "\n";
                    if (IsPhysicalDamage != null)
                        message += "Was there any Physical Damage: " + IsPhysicalDamage + "\n";
                    string email = GetEmail(HttpContext.Session.GetString("suserid"));
                    string retSupp = "";
                    string otheremail = "";
                    retSupp = GetEmailByGrp(SPKEY, "(All)", "D");
                    otheremail = GetEmailByGrp(SPKEY, SITEID, "S");
                    email = email + ";" + retSupp + ";" + otheremail;
                    SendEmail(email, subject, message);
                    
                    rtn = "Unacceptable";
                }
            }
            
            return rtn;
        }


        public string IPProcessed(string SITEID, int KITSET, string proctype, string RECVDDATE, string isExcrusion, string SPAPDTC, string Isapproved, string username, string IsPhysicalDamage)
        {
            string rtn = "";
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            //string SITEID = HttpContext.Session.GetString("sesCenter");
            string sql = "";
            if (isExcrusion == "Yes" && Isapproved == "Yes")
                sql = "UPDATE BIL_IP_RANGE SET  RECVDBY = @RECVDBY, RECVDDATE = @RECVDDATE, KITSTAT = '" + proctype + "', isExcrusion = @isExcrusion, SPAPDTC = @SPAPDTC, IsPhysicalDamage = @IsPhysicalDamage  WHERE KITSET = @KITSET AND KITSTAT = 'Shipped'";
            else if (isExcrusion == "No" && IsPhysicalDamage == "Yes")
                sql = "UPDATE BIL_IP_RANGE SET  RECVDBY = @RECVDBY, RECVDDATE = @RECVDDATE, KITSTAT = '" + proctype + "', isExcrusion = @isExcrusion, IsPhysicalDamage = @IsPhysicalDamage WHERE KITSET = @KITSET AND KITSTAT = 'Shipped'";
            else
            {
                rtn = "Error: Can not find a valid inventroy Type";
                return rtn;
            }
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            int rowsAfft;
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@KITSET", KITSET);
                cmd.Parameters.AddWithValue("@RECVDBY", userid);
                cmd.Parameters.AddWithValue("@isExcrusion", isExcrusion);
                cmd.Parameters.AddWithValue("@IsPhysicalDamage", IsPhysicalDamage);
                cmd.Parameters.AddWithValue("@RECVDDATE", RECVDDATE);
                if ((SPAPDTC != null) && (Isapproved == "Yes" && isExcrusion == "Yes"))
                {
                    cmd.Parameters.AddWithValue("@SPAPDTC", (object)SPAPDTC ?? DBNull.Value);
                }
                rowsAfft = (int)cmd.ExecuteNonQuery();

            }
            if (rowsAfft == 0)
            {
                rtn = "Error processing Kit at Site contact your PM/CRA.";
            }
            else
            {
                if (proctype == "Unacceptable")
                {
                    string UserName = GetUserName(HttpContext.Session.GetString("suserid"));
                    string subject = "Webview RTSM - Shipment Processed at Site " + SITEID + "";
                    string message = "Protocol: Webview RTSM - Test " + "\n" + "Shipment ID: " + KITSET + " has been received in Unacceptable condition by " + UserName + " at Site " + SITEID + "." + "\n";
                    message += "Received Date: " + RECVDDATE + "\n";
                    message += "Temperature Excursion: " + isExcrusion + "\n";
                    if (Isapproved != null)
                        message += "Sponsor approved the use of quarantined IP after temperature excursion?: " + Isapproved + "\n";
                    if (SPAPDTC != null)
                        message += "Date Of Sponsor Approval: " + SPAPDTC + "\n";
                    if (IsPhysicalDamage != null)
                        message += "Was there any Physical Damage: " + IsPhysicalDamage + "\n";
                    string email = GetEmail(HttpContext.Session.GetString("suserid"));
                    string retSupp = "";
                    string otheremail = "";
                    retSupp = GetEmailByGrp(SPKEY, "(All)", "D");
                    otheremail = GetEmailByGrp(SPKEY, SITEID, "S");
                    email = email + ";" + retSupp + ";" + otheremail;

                    SendEmail(email, subject, message);
                    
                    rtn = "Unacceptable";
                }
            }

            return rtn;
        }
        public int CountKitSet(int KITSET)
        {
            int retVal = 0;
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            string sqlState = "SELECT COUNT(*) FROM [BIL_IP_RANGE] WHERE (KITSET = " + KITSET + ") AND (KITSTAT = 'Shipped') ";
            SqlCommand cmd = new SqlCommand(sqlState, conn);
            using (conn)
            {
                conn.Open();
                int count = (int)cmd.ExecuteScalar();
                if (count > 0)
                {
                    retVal = count;
                }

            }
            conn.Close();
            return retVal;

        }
        public string UpdateRows(string SITEID, string SelectedKitKeys, int KITSET, string RECVDDATE, string isExcrusion, string SPAPDTC, string Isapproved, string username, int totalkit, string IsPhysicalDamage)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            //string SITEID = HttpContext.Session.GetString("sesCenter");
            string rtnVal = "";
            int rowsAfft = 0;
            string[] arr = SelectedKitKeys.Split(',');
            int[] Key = arr.Select(int.Parse).ToArray();
            int damageKits = 0;
            for (int i = 0; i < Key.Count(); i++)
            {
                InsertReceived(Key[i], KITSET, RECVDDATE, isExcrusion, SPAPDTC, Isapproved, username, IsPhysicalDamage);
            }
            if (totalkit != Key.Count())
            {
                
                SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
                string sql = "UPDATE BIL_IP_RANGE SET  RECVDBY = @RECVDBY, RECVDDATE = @RECVDDATE, KITSTAT = 'Unacceptable', isExcrusion = @isExcrusion, IsPhysicalDamage = @IsPhysicalDamage, SPAPDTC = @SPAPDTC  WHERE (KITSET = @KITSET) AND (KITSTAT = 'Shipped')";
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                SqlCommand cmd = new SqlCommand(sql, conn);
                
                using (conn)
                {
                    conn.Open();
                    cmd.Parameters.AddWithValue("@KITSET", KITSET);
                    cmd.Parameters.AddWithValue("@RECVDBY", userid);
                    cmd.Parameters.AddWithValue("@isExcrusion", isExcrusion);
                    cmd.Parameters.AddWithValue("@RECVDDATE", RECVDDATE);
                    cmd.Parameters.AddWithValue("@IsPhysicalDamage", IsPhysicalDamage);
                  //  if ((SPAPDTC != null) && (Isapproved == "Yes" && isExcrusion == "Yes"))
                    //{
                        cmd.Parameters.AddWithValue("@SPAPDTC", (object)SPAPDTC ?? DBNull.Value);
                    //}

                    rowsAfft = (int)cmd.ExecuteNonQuery();
                    if (rowsAfft == 0)
                        rtnVal = "Error processing Kit at Site contact your PM/CRA.";
                }
                

                conn.Close();
            }
            if(rowsAfft != 0)
            {
                damageKits = totalkit - Key.Count();
                string UserName = GetUserName(HttpContext.Session.GetString("suserid"));
                string subject = "Webview RTSM - Shipment Processed for [Amarex][Webview RTSM - Test] at Site " + SITEID + "";
                string message = "Protocol: Webview RTSM " + "\n" + "Shipment ID: " + KITSET + " has been processed by " + UserName + " at Site " + SITEID + "." + Environment.NewLine;
                message += "Kit Received in acceptable condition: " + Key.Count() + Environment.NewLine;
                message += "Kit Received in Unacceptable condition: " + damageKits + Environment.NewLine;
                message += "Received Date: " + RECVDDATE + "\n";
                message += "Temperature Excursion: " + isExcrusion + "\n";
                if (Isapproved != null)
                    message += "Sponsor approved the use of quarantined IP after temperature excursion?: " + Isapproved + "\n";
                if (SPAPDTC != null)
                    message += "Date Of Sponsor Approval: " + SPAPDTC + "\n";
                if (IsPhysicalDamage != null)
                    message += "Was there any Physical Damage: " + IsPhysicalDamage + "\n";
                string email = GetEmail(HttpContext.Session.GetString("suserid"));
                string retSupp = "";
                string otheremail = "";
                retSupp = GetEmailByGrp(SPKEY, "(All)", "D");
                otheremail = GetEmailByGrp(SPKEY, SITEID, "S");
                email = email + ";" + retSupp + ";" + otheremail;
                
                SendEmail(email, subject, message);
               
            }
            if(rtnVal == "")
            {
                rtnVal = "Kits has been processed";
            }
            
            return rtnVal;

        }

        public void InsertReceived(int Key, int KITSET, string RECVDDATE, string isExcrusion, string SPAPDTC, string Isapproved, string username, string IsPhysicalDamage)
        {
            string userid = HttpContext.Session.GetString("suserid");
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            string sql = "";
            if ((SPAPDTC != null) && (Isapproved == "Yes" && isExcrusion == "Yes"))
                sql = "UPDATE BIL_IP_RANGE SET RECVDBY = @RECVDBY, RECVDDATE = @RECVDDATE, KITSTAT = 'Acceptable', isExcrusion = @isExcrusion, IsPhysicalDamage = @IsPhysicalDamage, SPAPDTC = @SPAPDTC WHERE KITKEY = @KITKEY AND KITSTAT = 'Shipped'";
            else
                sql = "UPDATE BIL_IP_RANGE SET RECVDBY = @RECVDBY, RECVDDATE = @RECVDDATE, KITSTAT = 'Acceptable', isExcrusion = @isExcrusion, IsPhysicalDamage = @IsPhysicalDamage WHERE KITKEY = @KITKEY AND KITSTAT = 'Shipped'";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlCommand cmd = new SqlCommand(sql, conn);
            int rowsAfft;
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@KITKEY", Key);
                cmd.Parameters.AddWithValue("@RECVDBY", userid);
                cmd.Parameters.AddWithValue("@isExcrusion", isExcrusion);
                cmd.Parameters.AddWithValue("@RECVDDATE", RECVDDATE);
                if ((SPAPDTC != null) && (Isapproved == "Yes" && isExcrusion == "Yes"))
                {
                    cmd.Parameters.AddWithValue("@SPAPDTC", (object)SPAPDTC ?? DBNull.Value);
                }
                cmd.Parameters.AddWithValue("@IsPhysicalDamage", IsPhysicalDamage);
                rowsAfft = (int)cmd.ExecuteNonQuery();
                
            }
            

            conn.Close();
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

        public IActionResult UploadFiles(int KITSET, string SITEID)
        {

            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            ViewBag.SiteID = SITEID;
            ViewBag.KITSET = KITSET;

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            string sql = "SELECT KITKEY, KitNumber, SITEID, KITSTAT FROM BIL_IP_RANGE WHERE KITSET = @KITKEY";
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
                    selectCmd.Parameters.AddWithValue("@KITKEY", KITSET);

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
        public IActionResult UploadFiles(int KITSET, string SITEID, IFormFile PackageFile, IFormFile LogFile)
        {
            if (PackageFile != null && LogFile != null)
            {
                try
                {
                    string[] allowedExtensions = { ".pdf" };
                    var fileExtension = Path.GetExtension(PackageFile.FileName).ToLower();
                    var fileExtension2 = Path.GetExtension(LogFile.FileName).ToLower();

                    if (allowedExtensions.Contains(fileExtension) && allowedExtensions.Contains(fileExtension2))
                    {
                        var rtnVal = LogFileUP(KITSET, SITEID, LogFile);
                        var rtnVal2 = PackingFileUP(KITSET, SITEID, PackageFile);
                        if ((rtnVal == "OK") || (rtnVal2 == "OK"))
                        {
                            TempData["Message"] = "Temperature log and packing list have been succesfully uploaded.";
                            return RedirectToAction("ReceiveIPTabs");
                        }
                        else
                        {
                            TempData["ErrorMessage"] = "Error occured while processing the request";
                            return RedirectToAction("ReceiveIPTabs");
                        }

                    }
                    else
                    {
                        TempData["ErrorMessage"] = "Invalid file type. Only PDF files(.PDF) are permitted. Please ensure that you are uploading the file with the correct format.";
                    }
                }
                catch (Exception ex)

                {
                    TempData["ErrorMessage"] = "An error occurred while uploading the file: " + ex.Message;
                }
            }else if(PackageFile != null)
            {
                try
                {
                    string[] allowedExtensions = { ".pdf" };
                    var fileExtension = Path.GetExtension(PackageFile.FileName).ToLower();
                    if (allowedExtensions.Contains(fileExtension))
                    {
                        var rtnVal2 = PackingFileUP(KITSET, SITEID, PackageFile);
                        if (rtnVal2 == "OK")
                        {
                            TempData["Message"] = "Packing List has been succesfully uploaded.";
                            return RedirectToAction("ReceiveIPTabs");
                        }
                        else
                        {
                            TempData["ErrorMessage"] = "Error occured while processing the request";
                            return RedirectToAction("ReceiveIPTabs");
                        }

                    }
                    else
                    {
                        TempData["ErrorMessage"] = "Invalid file type. Only PDF files(.PDF) are permitted. Please ensure that you are uploading the file with the correct format.";
                    }
                }
                catch (Exception ex)

                {
                    TempData["ErrorMessage"] = "An error occurred while uploading the file: " + ex.Message;
                }
            }
            else if(LogFile != null)
            {
                try
                {
                    string[] allowedExtensions = { ".pdf" };
                    var fileExtension2 = Path.GetExtension(LogFile.FileName).ToLower();

                    if (allowedExtensions.Contains(fileExtension2))
                    {
                        var rtnVal = LogFileUP(KITSET, SITEID, LogFile);
                        if (rtnVal == "OK")
                        {
                            TempData["Message"] = "Temperature Log has been succesfully uploaded.";
                            return RedirectToAction("ReceiveIPTabs");
                        }
                        else
                        {
                            TempData["ErrorMessage"] = "Error occured while processing the request";
                            return RedirectToAction("ReceiveIPTabs");
                        }

                    }
                    else
                    {
                        TempData["ErrorMessage"] = "Invalid file type. Only PDF files(.PDF) are permitted. Please ensure that you are uploading the file with the correct format.";
                    }
                }
                catch (Exception ex)

                {
                    TempData["ErrorMessage"] = "An error occurred while uploading the file: " + ex.Message;
                }
            }
            else
            {
                TempData["ErrorMessage"] = "No file is available to upload.";
                
            }

            ViewBag.SiteID = SITEID;
            ViewBag.KITSET = KITSET;
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            string sql = "SELECT KITKEY, KitNumber, SITEID, KITSTAT FROM BIL_IP_RANGE WHERE KITSET = @KITKEY";
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
                    selectCmd.Parameters.AddWithValue("@KITKEY", KITSET);

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
        public string PackingFileUP(int KITSET, string SITEID, IFormFile PackageFile)
        {
            string rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            if (PackageFile != null)
            {
                string[] allowedExtensions = { ".pdf" };
                var fileExtension = Path.GetExtension(PackageFile.FileName).ToLower();
                if (allowedExtensions.Contains(fileExtension))
                {

                    Byte[] Packbytes = null;

                    using (MemoryStream ms = new MemoryStream())
                    {
                        PackageFile.OpenReadStream().CopyTo(ms);
                        Packbytes = ms.ToArray();
                    }

                    SqlCommand cmd = new SqlCommand("INSERT INTO IPUploads (SPKEY, SITEID, FILENAME, KITSET, TYPEUL, DataFile, ADDUSER, ADDDATE) values (@SPKEY, @SITEID, @FILENAME, @KITSET, @TYPEUL, @DataFile, @ADDUSER, @ADDDATE)", con);
                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.AddWithValue("@SPKEY", int.Parse(HttpContext.Session.GetString("sesSPKey")));
                    cmd.Parameters.AddWithValue("@FILENAME", PackageFile.FileName);
                    cmd.Parameters.AddWithValue("@DataFile", Packbytes);
                    cmd.Parameters.AddWithValue("@TYPEUL", "Packing Slip");
                    cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));
                    cmd.Parameters.AddWithValue("@SITEID", SITEID);
                    cmd.Parameters.AddWithValue("@KITSET", KITSET);
                    cmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                    con.Open();
                    ViewBag.PackID = cmd.ExecuteScalar();
                    con.Close();
                    rtnVal = "OK";
                }
                else
                    rtnVal = "Invalid";
            }
            return rtnVal;
        }

        public string LogFileUP(int KITSET, string SITEID, IFormFile LogFile)
        {
            string rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            if (LogFile != null)
            {
                string[] allowedExtensions = { ".pdf" };
                var fileExtension = Path.GetExtension(LogFile.FileName).ToLower();
                if (allowedExtensions.Contains(fileExtension))
                {
                    Byte[] Logbytes = null;

                    using (MemoryStream ms = new MemoryStream())
                    {
                        LogFile.OpenReadStream().CopyTo(ms);
                        Logbytes = ms.ToArray();
                    }

                    SqlCommand cmd2 = new SqlCommand("INSERT INTO IPUploads (SPKEY, SITEID, FILENAME, KITSET, TYPEUL, DataFile, ADDUSER, ADDDATE) values (@SPKEY, @SITEID, @FILENAME, @KITSET, @TYPEUL, @DataFile, @ADDUSER, @ADDDATE)", con);
                    cmd2.CommandType = CommandType.Text;

                    cmd2.Parameters.AddWithValue("@SPKEY", int.Parse(HttpContext.Session.GetString("sesSPKey")));
                    cmd2.Parameters.AddWithValue("@FILENAME", LogFile.FileName);
                    cmd2.Parameters.AddWithValue("@DataFile", Logbytes);
                    cmd2.Parameters.AddWithValue("@TYPEUL", "Temperature Log");
                    cmd2.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));
                    cmd2.Parameters.AddWithValue("@SITEID", SITEID);
                    cmd2.Parameters.AddWithValue("@KITSET", KITSET);
                    cmd2.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                    con.Open();
                    ViewBag.LogID = cmd2.ExecuteScalar();
                    con.Close();
                    rtnVal = "OK";
                }
                else
                    rtnVal = "Invalid";
            }
            return rtnVal;
        }

        public IActionResult DownloadFiles(string KITSET, string SITEID)
        {

            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            ViewBag.SiteID = SITEID;
            ViewBag.KITSET = KITSET;

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            string sql = "SELECT UPLOAD_KEY, SITEID, KITSET, FILENAME, KITSET, TYPEUL, ADDUSER, ADDDATE FROM IPUploads WHERE (KITSET = @KITSET) AND (SPKEY = " + SPKEY + ") AND ([FUSTATUS] = 'ACTIVE')";
            SqlCommand cmd = new SqlCommand(sql, conn);
            //ShipIP temp = new ShipIP();
            var list = new IPReporting();
            var IPFilesUP = new List<IPUploadFiles>();
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
                        var temp = new IPUploadFiles();
                        temp.UPLOAD_KEY = (int)rdr["UPLOAD_KEY"];
                        temp.FILENAME = rdr["FILENAME"].ToString();
                        temp.SITEID = rdr["SITEID"].ToString();
                        //temp.KITSET = (short)rdr["KITSET"];
                        temp.TYPEUL = rdr["TYPEUL"].ToString();
                        temp.ADDUSER = rdr["ADDUSER"].ToString();
                        temp.ADDDATE = (DateTime)rdr["ADDDATE"];



                        IPFilesUP.Add(temp);
                    }
                    rdr.Close();
                }
                con.Close();
            }

            list.IPFiles = IPFilesUP;
            return View(list);
        }


        public IActionResult Download(int UPLOAD_KEY)
        {
            var filetype = "xls";

            byte[] bytes = null;

            string name = string.Empty;


            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            using (SqlConnection con = new SqlConnection(connectionString))

            {

                using (SqlCommand cmd = new SqlCommand("Select * from [IPUploads] where UPLOAD_KEY = @id", con))

                {

                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.AddWithValue("@id", UPLOAD_KEY);

                    cmd.Connection = con;

                    con.Open();

                    SqlDataReader sdr = cmd.ExecuteReader();

                    sdr.Read();

                    if (sdr.HasRows)

                    {

                        bytes = (byte[])sdr["DataFile"];

                        filetype = "application/pdf";

                        name = Convert.ToString(sdr["FILENAME"]);

                    }

                    con.Close();

                }

            }
            TempData["DownloadComplete"] = "true";

            return File(bytes, filetype, name, lastModified: DateTime.UtcNow.AddSeconds(-5),

        entityTag: new Microsoft.Net.Http.Headers.EntityTagHeaderValue("\"MyCalculatedEtagValue\""));

        }


        /////////
        public string ChkSiteInv2(int spkey, string siteid)
        {
            string rtnVal = "";
            try
            {
                string sqlState;
                sqlState = "SELECT COUNT(*) AS CHKKITS FROM BIL_IP_RANGE WHERE (SITEID = '" + siteid + "') AND (TreatmentGroup IS NOT NULL) AND (ASSIGNED IS NULL) AND (KITSTAT = 'Acceptable') AND (RECVDBY IS NOT NULL)";
                SqlConnection con = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
                SqlCommand cmd = new SqlCommand(sqlState, con);
                con.Open();
                int cntArm = (int)cmd.ExecuteScalar();
                if (cntArm <= 8)
                {
                    string sqlState1 = "SELECT *, DATEDIFF(dd,BIL_IP_RANGE.SENDDATE,getdate()) AS SENDDAYS FROM BIL_IP_RANGE WHERE (SITEID = '" + siteid + "') AND (TreatmentGroup IS NOT NULL) AND (KITSTAT = 'Shipped')";
                    SqlCommand cmd2 = new SqlCommand(sqlState1, con);
                    SqlDataReader reader = cmd2.ExecuteReader();
                    if (reader.Read())
                    {
                        if (Convert.ToInt16(reader["SENDDAYS"]) > 4)
                        {
                            string msgBody;
                            msgBody = "For Webview RTSM - Site inventory low, Kit in Shipped status for over 4 days " + Environment.NewLine;
                            msgBody += "Site: " + siteid + Environment.NewLine + "Kit Set: " + reader["KITSET"].ToString();
                            string retSupp = "";
                            //retSupp = GetEmailByGrp(spkey, "(All)", "D");
                            if (retSupp == "")
                            {
                                SendEmail("sidran@amarexcro.com", "Webview RTSM - Kit Shipment - Site " + siteid + " - Low Inv with Shipped status - No Supp", msgBody);
                            }
                            else
                            {
                                SendEmail(retSupp + ";" + "sidran@amarexcro.com", "Webview RTSM - Kit Shipment - Site " + siteid + " - Low Inv with Shipped status", msgBody);
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
                        sqlState2 += "(SELECT TOP 8 KITKEY FROM BIL_IP_RANGE WHERE (TreatmentGroup IS NOT NULL) AND ([ATDEPOT] is not NULL) AND (SITEID IS NULL) ORDER BY KitNumber); ";
                        SqlCommand cmd3 = new SqlCommand(sqlState2, con);
                        int rowsAfft = (int)cmd3.ExecuteNonQuery();
                        if (rowsAfft <= 0)
                        {
                            SendEmail("sidran@amarexcro.com", "Webview RTSM - Auto Re-supply - Site " + siteid + " -  Unable to find kits", "Unable to find kits");
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
                                SendEmail("sidran@amarexcro.com", "Webview RTSM - Auto Re-supply - Site " + siteid + " -  Shipment process, unable to find kits", "Shipment process, unable to find kits ");
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

        public IActionResult ViewKits(int KITSET, string SITEID)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            ViewBag.SiteID = SITEID;
            ViewBag.KITSET = KITSET;

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            string sql = "SELECT KITKEY, KitNumber, SITEID, KITSTAT FROM BIL_IP_RANGE WHERE KITSET = @KITSET";
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


        public void LogActivity(int KITSET)
        {
           
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql = "INSERT INTO Log_Activity (SPKEY, USERID, LOG_DESC,  LOG_TYPE, KITSET, SITEID) VALUES (@SPKEY, @USERID, @LOG_DESC,  @LOG_TYPE, @KITSET, @SITEID)";
            SqlConnection conn = new SqlConnection(connectionString);

            using (conn)
            {
                conn.Open();
                
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@SPKEY", SPKEY);
                    cmd.Parameters.AddWithValue("@USERID", userid);
                    cmd.Parameters.AddWithValue("@LOG_DESC", "Kit Processed");
                    cmd.Parameters.AddWithValue("@LOG_TYPE", "IP Receive");
                    cmd.Parameters.AddWithValue("@KITSET", KITSET);
                    cmd.Parameters.AddWithValue("@SITEID", SITEID);
                    cmd.ExecuteNonQuery();
               
            }
            

        }


    }
}












