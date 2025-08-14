using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using Webview_IRT.Models;

namespace RTSM_OLSingleArm.Controllers
{
    public class IPDepotController : Controller
    {
        public string connectionString;
        readonly IConfiguration _configuration;
        public IPDepotController(IConfiguration configuration)
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
        public IActionResult IPDepotHome()
        {
            return View();
        }
        public IActionResult IPDepotReg()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql = @"
                        SELECT 
        ship.Courier, 
        ship.TrackNo, 
        ship.ExpiryDate, 
        ship.LotNo, 
        ship.IPLblShipDTC, 
        ship.ILSKEY, 
        ship.ADDUSER,
        STUFF((
            SELECT ', ' + CONCAT(req.RangeStr, ' - ', req.RangeEnd)
            FROM IP_LABEL_SHIP_REQ AS req
            WHERE ship.ILSKEY = req.ILSKEY
            FOR XML PATH('')), 1, 2, '') AS RangeStrEnd,
        STUFF((
            SELECT ', ' + CONCAT(depot.RangeStr, ' - ', depot.RangeEnd, ' $ ', depot.IPDepotDTC, ' $ ', depot.ADDUSER, '')
            FROM IP_LABEL_DEPOT_REL AS depot
            WHERE ship.ILSKEY = depot.ILSKEY
            FOR XML PATH('')), 1, 2, '') AS DepotRangeStrEnd
    FROM 
        IP_LABEL_SHIP AS ship
    LEFT JOIN 
        IP_LABEL_DEPOT_REL AS depot ON ship.ILSKEY = depot.ILSKEY
    WHERE 
        ship.IPSTAT IS NULL
    GROUP BY 
        ship.Courier, 
        ship.TrackNo, 
        ship.ExpiryDate, 
        ship.LotNo, 
        ship.IPLblShipDTC, 
        ship.ILSKEY, 
        ship.ADDUSER
    ORDER BY 
        ship.ILSKEY;
            ";
            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            var list = new IPlabel();
            var ship = new List<IPShip>();
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var temp = new IPShip();
                    temp.ILSKEY = (int)rdr["ILSKEY"];
                    temp.TrackNo = rdr["TrackNo"].ToString();
                    temp.Courier = rdr["Courier"].ToString();
                    temp.ExpiryDate = rdr["ExpiryDate"].ToString();
                    temp.LotNo = rdr["LotNo"].ToString();
                    temp.ADDUSER = rdr["ADDUSER"].ToString();
                    temp.IPLblShipDTC = rdr["IPLblShipDTC"].ToString();
                    temp.RangeStrEnd = rdr["RangeStrEnd"].ToString();
                    temp.DepotRangeStrEnd = rdr["DepotRangeStrEnd"].ToString();

                    ship.Add(temp);
                }
                rdr.Close();


            }
            con.Close();
            list.IPShip = ship;
            return View(list);


        }
        [HttpPost]
        public IActionResult IPDepotReg(IPlabel Request, string username, string password)
        {
            string userid = HttpContext.Session.GetString("suserid");
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql = @"
                    SELECT 
                ship.Courier, 
                ship.TrackNo, 
                ship.ExpiryDate, 
                ship.LotNo, 
                ship.IPLblShipDTC, 
                ship.ILSKEY, 
                ship.ADDUSER,
                STUFF((
                    SELECT ', ' + CONCAT(req.RangeStr, ' - ', req.RangeEnd)
                    FROM IP_LABEL_SHIP_REQ AS req
                    WHERE ship.ILSKEY = req.ILSKEY
                    FOR XML PATH('')), 1, 2, '') AS RangeStrEnd,
                STUFF((
                    SELECT ', ' + CONCAT(depot.RangeStr, ' - ', depot.RangeEnd, ' $ ', depot.IPDepotDTC, ' $ ', depot.ADDUSER, '')
                    FROM IP_LABEL_DEPOT_REL AS depot
                    WHERE ship.ILSKEY = depot.ILSKEY
                    FOR XML PATH('')), 1, 2, '') AS DepotRangeStrEnd
            FROM 
                IP_LABEL_SHIP AS ship
            LEFT JOIN 
                IP_LABEL_DEPOT_REL AS depot ON ship.ILSKEY = depot.ILSKEY
            WHERE 
                ship.IPSTAT IS NULL
            GROUP BY 
                ship.Courier, 
                ship.TrackNo, 
                ship.ExpiryDate, 
                ship.LotNo, 
                ship.IPLblShipDTC, 
                ship.ILSKEY, 
                ship.ADDUSER
            ORDER BY 
                ship.ILSKEY;
            ";
            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            var list = new IPlabel();
            var ship = new List<IPShip>();
            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var temp = new IPShip();
                    temp.ILSKEY = (int)rdr["ILSKEY"];
                    temp.TrackNo = rdr["TrackNo"].ToString();
                    temp.Courier = rdr["Courier"].ToString();
                    temp.ExpiryDate = rdr["ExpiryDate"].ToString();
                    temp.LotNo = rdr["LotNo"].ToString();
                    temp.ADDUSER = rdr["ADDUSER"].ToString();
                    temp.IPLblShipDTC = rdr["IPLblShipDTC"].ToString();
                    temp.RangeStrEnd = rdr["RangeStrEnd"].ToString();
                    temp.DepotRangeStrEnd = rdr["DepotRangeStrEnd"].ToString();
                    ship.Add(temp);
                }
                rdr.Close();


            }
            con.Close();
            list.IPShip = ship;

            if (IsValidUser(username, password) && (string.Equals(userid, username, StringComparison.OrdinalIgnoreCase) || userid == "sidran"))
            {
                var chkStr = "";
                var chkEnd = "";
                int intkStr = 0;
                int intEnd = 0;
                var rtnVal = "";
                string[] startIPArray = Request.StartIP;
                string[] endIPArray = Request.EndIP;
                int previousStart = 0; // Variable to store the previous start value
                int previousEnd = 0; // Variable to store the previous end value
                HashSet<string> seenRanges = new HashSet<string>(); // HashSet to store seen IP ranges

                for (int i = 0; i < startIPArray.Length; i++)
                {
                    if ((startIPArray[i].Length != 5) || (endIPArray[i].Length != 5))
                    {
                        TempData["ErrorMessage"] = "Kit number format is incorrect.";
                        return View(list);
                    }
                    chkStr = startIPArray[i].Substring(1, 4);
                    chkEnd = endIPArray[i].Substring(1, 4);
                    if (!int.TryParse(chkStr, out int result) || !int.TryParse(chkEnd, out int result1))
                    {
                        TempData["ErrorMessage"] = "Kit number format is incorrect.";
                        return View(list);
                    }
                    else
                    {
                        intkStr = result;
                        intEnd = result1;
                    }
                    //string range = startIPArray[i] + "-" + endIPArray[i];
                    //if (seenRanges.Contains(range))
                    //{
                    //    TempData["ErrorMessage"] = "Part of the entered kit range is duplicate. Please verify and update.";
                    //    return View(list);
                    //}
                    //seenRanges.Add(range);
                    //if ((intkStr > 0) && (intEnd > 0))
                    //{
                    //    if (intkStr > intEnd)
                    //    {
                    //        TempData["ErrorMessage"] = "Start Kit Number cannot be greater then End.";
                    //        return View(list);
                    //    }

                    //    if (i > 0 && intkStr >= previousStart && intkStr <= previousEnd)
                    //    {
                    //        TempData["ErrorMessage"] = "Part of the entered kit range is duplicate. Please verify and update.";
                    //        return View(list);
                    //    }

                    //    previousStart = intkStr; // Update previous start value for the next iteration
                    //    previousEnd = intEnd;    // Update previous end value for the next iteration

                    //}

                    if (intkStr > 0 && intEnd > 0)
                    {
                        if (intkStr > intEnd)
                        {
                            TempData["ErrorMessage"] = "Start Kit Number cannot be greater than End.";
                            return View(list);
                        }

                        string range = startIPArray[i] + "-" + endIPArray[i]; // Construct range string
                        foreach (string seenRange in seenRanges)
                        {
                            if (CheckFullRangeOverlap(range, seenRange))
                            {
                                TempData["ErrorMessage"] = "Part of the entered kit range is duplicate. Please verify and update.";
                                return View(list);
                            }
                        }
                        seenRanges.Add(range);

                        if (CheckRangeOverlap(intkStr, intEnd, previousStart, previousEnd))
                        {
                            TempData["ErrorMessage"] = "Part of the entered kit range is duplicate. Please verify and update.";
                            return View(list);
                        }
                        previousStart = intkStr;
                        previousEnd = intEnd;
                    }


                    rtnVal = CheckBox(startIPArray[i], endIPArray[i]);
                    if (rtnVal == "")
                    {

                        TempData["ErrorMessage"] = "Part of the entered Kit range is not valid. Please verify and update.";
                        return View(list);
                    }

                    //Check if IP is already at depot
                    rtnVal = ChkToDepotIPDepot(startIPArray[i], endIPArray[i]);
                    if (rtnVal != "")
                    {
                        //TempData["ErrorMessage"] = "Part of the IP/Kit Range: " + startIPArray[i] + " - " + endIPArray[i] + "  have already been confirmed as received by Depot and thus will not be added to the Depot inventory again.";
                        TempData["ErrorMessage"] = "Part of the Kit Range have already been confirmed as received by Depot and thus will not be added to the Depot inventory again.";
                        return View(list);
                    }
                    //Check whether IP is Released to the depot or not
                    rtnVal = ChkRelToDepotIP(startIPArray[i], endIPArray[i]);
                    if (rtnVal != "")
                    {
                        TempData["ErrorMessage"] = "Part of the kit range have not been shipped or logged by the IP labeling company. Please (i) correct the kit range if entered incorrectly, and/or (ii) contact the PM/labeling company to correct the kit range information for kits received at depot but not logged by labeling company.";
                        return View(list);
                    }
                    //Validate Expiry date
                    string rtnExpiry = ChkToDepotIPExpr(startIPArray[i], endIPArray[i], Request.IPLblShipExpiryDate);
                    //Validate Lot No
                    string rtnLot = ChkToDepotIPLot(startIPArray[i], endIPArray[i], Request.IPLblShipLotNo);

                    //Validate Expiry date & Lot No
                    if ((rtnExpiry != "") && (rtnLot != ""))
                    {

                        TempData["ErrorMessage"] = "Expiry Date & Lot Number do not match the details provided by the IP labeling company.";
                        return View(list);
                    }
                    //Validate Expiry date
                    if (rtnExpiry != "")
                    {
                        TempData["ErrorMessage"] = "Expiry date does not match the details provided by the IP labeling company.";
                        return View(list);
                    }
                    //Validate Lot No
                    if (rtnLot != "")
                    {
                        TempData["ErrorMessage"] = "Lot number does not match the details provided by the IP labeling company.";
                        return View(list);
                    }
                }

                var ILSKEY = GetILSKEY(startIPArray[0]);

                //Validate the Received date must be after IP shipment Date
                DateTime Shipmentdate = DateTime.Parse(ChkShipmentDate(ILSKEY));
                DateTime RecvDate = DateTime.Parse(Request.IPDepotDTC);

                if (Shipmentdate > RecvDate)
                {
                    TempData["ErrorMessage"] = "Received date should be on or after kit Shipment date";
                    return View(list);
                }

                DateTime IPLblShipExpiryDate = DateTime.Parse(Request.IPLblShipExpiryDate);

                // Add 6 months to the RecvDate
                DateTime sixMonthsLater = RecvDate.AddMonths(6);

                // Check if IPLblShipExpiryDate is less than six months later than RecvDate
                if (IPLblShipExpiryDate < sixMonthsLater)
                {
                    //TempData["ErrorMessage"] = "Expiry date must be at least 6 months from the received date.";
                    TempData["ErrorMessage"] = "Please ensure that the expiration date is at least six months from the received date. The current expiration date does not meet this requirement, so the shipment cannot be received in good condition. Please quarantine the shipment.";
                    return View(list);
                }


                var msgBody = "";
                var rtnValue = "";
                var message = "";
                
                var details = GetCourier(ILSKEY);
                string[] arr = details.Split(";");
                var Courier = arr[0];
                var TrackNo = arr[1];
                var IPLblShipDTC = arr[2];
                int count = 0;

                for (int i = 0; i < startIPArray.Length; i++)
                {
                    rtnValue = UpdtToDepotIP(startIPArray[i], endIPArray[i], Request.IPLblShipExpiryDate, Request.IPLblShipLotNo, userid, ILSKEY, Courier, TrackNo, IPLblShipDTC, Request.IPDepotDTC);
                    if ((rtnValue == "IP Label Error") || (rtnValue == "IP_LABEL_DEPOT_REL INSERT"))
                    {
                        TempData["ErrorMessage"] = rtnValue;
                        return View(list);
                    }
                    else
                    {
                        ///select
                        count += RowsAffect(startIPArray[i], endIPArray[i]);
                        message += startIPArray[i] + " to " + endIPArray[i];
                        if (i < startIPArray.Length - 1)
                        {
                            message += ", ";
                        }

                    }


                }
                message = "A total of " + count + " kit(s) released to the depot inventory. Range(s) from " + message + ".";
                TempData["Message"] = message;
                string randemail = GetProfVpe(SPKEY, "PMIPLblRel") + ";" + GetUserEmail() + ";" + GetProfVpe(SPKEY, "IPDepot");
                string subject = StudyName() + " - WebView RTSM - Depot kit release confirmation";
               
                msgBody += Environment.NewLine + "Protocol: " + StudyName() + "\n";
                msgBody += Environment.NewLine + "This email is to notify you that the following kit(s) have been released to the depot inventory, details below." + "\n";
                msgBody += Environment.NewLine + "Expiry Date: " + Request.IPLblShipExpiryDate;
                msgBody += Environment.NewLine + "Lot Number: " + Request.IPLblShipLotNo;
                msgBody += Environment.NewLine + message;
               
                SendEmail(randemail + ";" + "sidran@amarexcro.com", subject, msgBody);
                string sql2 = @"
                    SELECT 
                ship.Courier, 
                ship.TrackNo, 
                ship.ExpiryDate, 
                ship.LotNo, 
                ship.IPLblShipDTC, 
                ship.ILSKEY, 
                ship.ADDUSER,
                STUFF((
                    SELECT ', ' + CONCAT(req.RangeStr, ' - ', req.RangeEnd)
                    FROM IP_LABEL_SHIP_REQ AS req
                    WHERE ship.ILSKEY = req.ILSKEY
                    FOR XML PATH('')), 1, 2, '') AS RangeStrEnd,
                STUFF((
                    SELECT ', ' + CONCAT(depot.RangeStr, ' - ', depot.RangeEnd, ' $ ', depot.IPDepotDTC, ' $ ', depot.ADDUSER, '')
                    FROM IP_LABEL_DEPOT_REL AS depot
                    WHERE ship.ILSKEY = depot.ILSKEY
                    FOR XML PATH('')), 1, 2, '') AS DepotRangeStrEnd
            FROM 
                IP_LABEL_SHIP AS ship
            LEFT JOIN 
                IP_LABEL_DEPOT_REL AS depot ON ship.ILSKEY = depot.ILSKEY
            WHERE 
                ship.IPSTAT IS NULL
            GROUP BY 
                ship.Courier, 
                ship.TrackNo, 
                ship.ExpiryDate, 
                ship.LotNo, 
                ship.IPLblShipDTC, 
                ship.ILSKEY, 
                ship.ADDUSER
            ORDER BY 
                ship.ILSKEY;
            ";
                SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
                SqlCommand cmd2 = new SqlCommand(sql2, conn);
                var list2 = new IPlabel();
                var ship2 = new List<IPShip>();
                using (conn)
                {
                    conn.Open();
                    SqlDataReader rdr2 = cmd2.ExecuteReader();
                    while (rdr2.Read())
                    {
                        var temp = new IPShip();
                        temp.ILSKEY = (int)rdr2["ILSKEY"];
                        temp.TrackNo = rdr2["TrackNo"].ToString();
                        temp.Courier = rdr2["Courier"].ToString();
                        temp.ExpiryDate = rdr2["ExpiryDate"].ToString();
                        temp.LotNo = rdr2["LotNo"].ToString();
                        temp.ADDUSER = rdr2["ADDUSER"].ToString();
                        temp.IPLblShipDTC = rdr2["IPLblShipDTC"].ToString();
                        temp.RangeStrEnd = rdr2["RangeStrEnd"].ToString();
                        temp.DepotRangeStrEnd = rdr2["DepotRangeStrEnd"].ToString();
                        ship2.Add(temp);
                    }
                    rdr2.Close();


                }
                conn.Close();
                list2.IPShip = ship2;
                return View(list2);
            }
            else
            {
                // Invalid username or password
                string errorMessage = "Invalid username or password.";
                //ViewBag.ErrorMessage = errorMessage;
                TempData["ErrorMessage"] = errorMessage ;
                return View(list);
            }
        }

        public string CheckBox(string strIP, string endIP)
        {
            string rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new System.Data.SqlClient.SqlConnection(connectionString);
            string sql = "SELECT COUNT(*) FROM BIL_IP_RANGE WHERE KitNumber = '" + strIP + "' OR  KitNumber = '" + endIP + "'";
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                int count = (int)cmd.ExecuteScalar();
                if ((strIP == endIP && count == 1) || (count >= 2))
                {
                    // Handle invalid subjid
                    rtnVal = "Box Number is valid";

                }
                conn.Close();
            }
            return rtnVal;

        }

        public bool CheckRangeOverlap(int start1, int end1, int start2, int end2)
        {
            // Ranges overlap if one starts before the other ends, and vice versa
            return (start1 <= end2 && end1 >= start2);
        }

        public bool CheckFullRangeOverlap(string range1, string range2)
        {
            // Extract start and end values from range1 and range2
            int start1 = int.Parse(range1.Substring(2, 3));
            int end1 = int.Parse(range1.Substring(8, 3));
            int start2 = int.Parse(range2.Substring(2, 3));
            int end2 = int.Parse(range2.Substring(8, 3));

            // Check if ranges overlap using CheckRangeOverlap method
            return CheckRangeOverlap(start1, end1, start2, end2);
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

        public string checkidpwd(string username, string password)
        {

            var rtnVal = "";
            SecSSO chkSSO2 = new SecSSO();
            rtnVal = chkSSO2.ChkIDPWSSO(username, password, HttpContext.Session.GetString("sesuriSSIS"), HttpContext.Session.GetString("sesinstanceID"), HttpContext.Session.GetString("sesSecurityKey"), HttpContext.Session.GetString("sesAmarexDb"));
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
            conn.Close();
            return email;
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

        public string ChkToDepotIPExpr(string strIP, string endIP, string exprDT)
        {
            var rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();
                var sqlState = "";
                sqlState = "SELECT KitNumber FROM BIL_IP_RANGE  WHERE (([IPLblShipExpiryDate] ! = '" + exprDT + "') OR ([IPLblShipExpiryDate] is null)) ";
                sqlState += "AND ([KitNumber] between '" + strIP + "' AND '" + endIP + "')";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                   while (rdr.Read())
                   {
                      if (rtnVal == "")
                         {
                           rtnVal = rdr["KitNumber"].ToString();
                         }
                      else
                        {
                          rtnVal += "<br />" + rdr["KitNumber"].ToString();
                         }
                   }
                
            }
            conn.Close();
            return rtnVal;
        }

        public string ChkToDepotIPLot(string strIP, string endIP, string lot)
        {
            var rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();
                var sqlState = "";
               sqlState = "SELECT KitNumber FROM BIL_IP_RANGE  WHERE (([IPLblShipLotNo] ! = '" + lot + "') OR ([IPLblShipLotNo] is null)) ";
               sqlState += "AND ([KitNumber] between '" + strIP + "' AND '" + endIP + "')";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rtnVal == "")
                    {
                      rtnVal = rdr["KitNumber"].ToString();
                    }
                    else
                    {
                      rtnVal += "<br />" + rdr["KitNumber"].ToString();
                    }
                }
               

            }
            conn.Close();
            return rtnVal;
        }

        public string ChkRelToDepotIP(string strIP, string endIP)
        {
            var rtnVal = "";
            
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();
                var sqlState = "";
                sqlState = "SELECT KitNumber FROM BIL_IP_RANGE  WHERE ([ILSKEY] is NULL) ";
                sqlState += "AND ([KitNumber] between '" + strIP + "' AND '" + endIP + "')";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                   
                   if (rtnVal == "")
                   {
                      rtnVal = rdr["KitNumber"].ToString();
                   }
                   else
                   {
                      rtnVal += "<br />" + rdr["KitNumber"].ToString();
                   }
                        
                    
                }
                
            }
            conn.Close();
            return rtnVal;
        }

        public string ChkToDepotIPDepot(string strIP, string endIP)
        {
            var rtnVal = "";

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();
                var sqlState = "";
                sqlState = "SELECT KitNumber FROM BIL_IP_RANGE  WHERE ([ATDEPOT] is not NULL) ";
                sqlState += "AND ([KitNumber] between '" + strIP + "' AND '" + endIP + "')";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    if (rtnVal == "")
                    {
                        rtnVal = rdr["KitNumber"].ToString();
                    }
                    else
                    {
                        rtnVal += "<br />" + rdr["KitNumber"].ToString();
                    }
                }
                
            }
            conn.Close();
            return rtnVal;
        }

        public string UpdtToDepotIP(string strIP, string endIP, string expiry, string lotNo, string uid, int ILSKEY, string Courier, string TrackNo, string IPLblShipDTC, string IPDepotDTC)
        {
            string rtnVal = "IP Label Error";
            rtnVal = InsIPDepotRel(expiry, lotNo, strIP, endIP, uid, ILSKEY, Courier, TrackNo, IPLblShipDTC, IPDepotDTC);
            if (rtnVal != "OK")
            {
                return rtnVal;
            }

            string connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                
                    string sqlState = "";
                    //sqlState = "UPDATE BIL_IP_RANGE SET [ATDEPOT] = 'Yes', [ATDEPOTID] = '" + uid + "', [ATDEPOTDTC] = sysdatetime() WHERE ([SITEID] is null) AND ([PMShipToDepot] = 'Approve') AND ([ATDEPOT] is null) ";
                    sqlState = "UPDATE BIL_IP_RANGE SET [ATDEPOT] = @ATDEPOT, [ATDEPOTID] = @ATDEPOTID, [ATDEPOTDTC] = @ATDEPOTDTC , IPDepotDTC = @IPDepotDTC WHERE ([SITEID] is null) AND ([ATDEPOT] is null) ";
                    sqlState += "AND ([KitNumber] between '" + strIP + "' AND '" + endIP + "')";
                    SqlCommand cmd = new SqlCommand(sqlState, conn);
                //cmd.ExecuteNonQuery();
                cmd.Parameters.AddWithValue("@ATDEPOT", "Yes");
                cmd.Parameters.AddWithValue("@ATDEPOTID", uid);
                cmd.Parameters.AddWithValue("@ATDEPOTDTC", DateTime.Now);
                cmd.Parameters.AddWithValue("@IPDepotDTC", IPDepotDTC);


                int rowsAfft = (int)cmd.ExecuteNonQuery();
                    rtnVal = rowsAfft + " kit(s) released to the depot inventory.";
                
            }
            return rtnVal;
        }

        public string InsIPDepotRel(string expiry, string lotNo, string rangestr, string rangeend, string uid, int ILSKEY, string Courier, string TrackNo, string IPLblShipDTC, string IPDepotDTC)
        {
            string rtnVal = "OK";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();

                string sqlState = "";
                sqlState = "INSERT INTO [IP_LABEL_DEPOT_REL] ([ExpiryDate], [LotNo], IPLabelShip, RangeStr, RangeEnd, [ADDUSER], ILSKEY, Courier, TrackNo, IPLblShipDTC, IPDepotDTC) VALUES ('" + expiry + "', '" + lotNo + "', 'YES', '" + rangestr + "', '" + rangeend + "', '" + uid + "', '" + ILSKEY + "', '" + Courier + "', '" + TrackNo + "', '" + IPLblShipDTC + "', '" + IPDepotDTC + "')";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                int rowsAfft = (int)cmd.ExecuteNonQuery();
                if (rowsAfft <= 0)
                {
                    rtnVal = "IP_LABEL_DEPOT_REL INSERT";
                }

            }
            conn.Close();
            return rtnVal;
        }
        //Checked whether All Kits being processed by IP Depot
        public string Shipmentcompleted(string KitNumber, string userid)
        {
            var rtnVal = "";
            int ILSKEY = GetILSKEY(KitNumber);
            int count = 1;
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();

                string sqlState = "SELECT COUNT(*) FROM BIL_IP_RANGE WHERE [ATDEPOT] IS NULL AND ILSKEY = " + ILSKEY + "";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                count = (int)cmd.ExecuteScalar();

            }
            conn.Close();
            if(count == 0)
            {
                rtnVal = Updtcomplete(ILSKEY, userid);
            }
            return rtnVal;
            
        }

        public int GetILSKEY(string KitNumber)
        {
            var rtnVal = 0;

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();
                
                string sqlState = "SELECT ILSKEY FROM BIL_IP_RANGE  WHERE KitNumber = '" + KitNumber + "'";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    rtnVal = (int)rdr["ILSKEY"];
                }

            }
            conn.Close();
            return rtnVal;
        }

        public string Updtcomplete(int ilskey, string uid)
        {
            string rtnVal = "IP Label Update Error";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {

                string sqlState = "UPDATE BIL_IP_RANGE SET [ILSKEY] = null, [IPLblShipDTC] = null, [Courier] = null, [TrackNo] = null, [IPLblShipExpiryDate] = null, [IPLblShipLotNo] = null, [IPLblShipSysDTC] = null, [IPLblShipUID] = null WHERE ([ILSKEY] = " + ilskey + "); ";
                string sqlState2 = "UPDATE IP_LABEL_SHIP SET IP_LABEL_SHIP.CHANGEUSER = '" + uid + "', [CHANGEDATE] = sysdatetime(), [IPSTAT] = 'Completed' WHERE ([ILSKEY] = " + ilskey + ")";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                SqlCommand cmd2 = new SqlCommand(sqlState2, conn);
                conn.Open();
                cmd.ExecuteNonQuery();
                cmd2.ExecuteNonQuery();
                rtnVal = "OK";
            }
            conn.Close();
            return rtnVal;
        }

        public string GetCourier(int ILSKEY)
        {
            var rtnVal = "";

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();

                string sqlState = "SELECT Courier, TrackNo, IPLblShipDTC  FROM IP_LABEL_SHIP WHERE ILSKEY = " + ILSKEY + "";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    rtnVal = rdr["Courier"].ToString() + ";" ;
                    rtnVal += rdr["TrackNo"].ToString() + ";";
                    rtnVal += rdr["IPLblShipDTC"].ToString() + ";";

                }

            }
            conn.Close();
            return rtnVal;
        }

       public string ChkShipmentDate(int ILSKEY)
        {
            var rtnVal = "";

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();

                string sqlState = "SELECT IPLblShipDTC FROM IP_LABEL_SHIP WHERE ILSKEY = " + ILSKEY + "";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    rtnVal = rdr["IPLblShipDTC"].ToString();

                }

            }
            conn.Close();
            return rtnVal;
        }


        public int RowsAffect(string startIP, string endIP)
        {
            int count = 0;
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();

                string sqlState = "SELECT COUNT(*) FROM BIL_IP_RANGE WHERE ATDEPOT = 'Yes' AND ([KitNumber] BETWEEN '" + startIP + "' AND '" + endIP + "')";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                count = (int)cmd.ExecuteScalar();

            }
            conn.Close();
            return count;
        }

    }
}
