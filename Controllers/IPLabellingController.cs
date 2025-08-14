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
    public class IPLabellingController : Controller
    {

        public string connectionString;
        readonly IConfiguration _configuration;
        public IPLabellingController(IConfiguration configuration)
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
        public IActionResult IPLabellingHome()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
   
            string sql = "SELECT UPLOAD_KEY, UPLOAD_DESC, FILENAME, EMAILSENT, TYPEUL, ADDUSER, ADDDATE FROM IP_BIO_UPLOADS WHERE TYPEUL = 'Kit' AND IsHide is null AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY  FILENAME, UPLOAD_KEY , ADDDATE ";
            //SelectCommand="SELECT [Courier], [TrackNo], [ExpiryDate], [LotNo], [IPLblShipDTC], [ILSKEY] FROM [IP_LABEL_SHIP] WHERE [IPSTAT] is null ORDER BY [ILSKEY]
            //string sql2 = "SELECT Courier, TrackNo, ExpiryDate, LotNo, IPLblShipDTC, ILSKEY, ADDUSER FROM IP_LABEL_SHIP WHERE IPSTAT is null ORDER BY ILSKEY";
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
                    SELECT ', ' + CONCAT(depot.RangeStr, ' - ', depot.RangeEnd, ' $ ', depot.IPDepotDTC, ' $ ', depot.ADDUSER,'')
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
                ship.ILSKEY;";

            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlCommand cmd2 = new SqlCommand(sql2, con);
            //SqlCommand cmd3 = new SqlCommand(sql3, con);

            var list = new IPLabelList();
            var kit = new List<KitUpload>();
            var ship = new List<IPShip>();
            //var rand = new List<RandUpload>();
            //var remove = new List<RemovedKits>();


            using (con)
            {
                con.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var temp = new KitUpload();
                    temp.UPLOAD_KEY = (int)rdr["UPLOAD_KEY"];
                    temp.UPLOAD_DESC = rdr["UPLOAD_DESC"].ToString();
                    temp.FILENAME = rdr["FILENAME"].ToString();
                    temp.EMAILSENT = rdr["EMAILSENT"].ToString();
                    temp.TYPEUL = rdr["TYPEUL"].ToString();
                    temp.ADDUSER = rdr["ADDUSER"].ToString();
                    temp.ADDDATE = (DateTime)rdr["ADDDATE"];
                    kit.Add(temp);
                }
                rdr.Close();
                SqlDataReader rdr3 = cmd2.ExecuteReader();
                while (rdr3.Read())
                {
                    var temp = new IPShip();
                    temp.ILSKEY = (int)rdr3["ILSKEY"];
                    temp.TrackNo = rdr3["TrackNo"].ToString();
                    temp.Courier = rdr3["Courier"].ToString();
                    temp.ExpiryDate = rdr3["ExpiryDate"].ToString();
                    temp.LotNo = rdr3["LotNo"].ToString();
                    temp.ADDUSER = rdr3["ADDUSER"].ToString();
                    temp.IPLblShipDTC = rdr3["IPLblShipDTC"].ToString();
                    temp.RangeStrEnd = rdr3["RangeStrEnd"].ToString();
                    temp.DepotRangeStrEnd = rdr3["DepotRangeStrEnd"].ToString();
                    ship.Add(temp);
                }
                rdr3.Close();
                

            }
            con.Close();
            list.kitupload = kit;
            list.IPShip = ship;
            //list.removed = remove;
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

                using (SqlCommand cmd = new SqlCommand("Select * from [IP_BIO_UPLOADS] where UPLOAD_KEY = @id", con))

                {

                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.AddWithValue("@id", UPLOAD_KEY);

                    cmd.Connection = con;

                    con.Open();

                    SqlDataReader sdr = cmd.ExecuteReader();

                    sdr.Read();

                    if (sdr.HasRows)

                    {

                        bytes = (byte[])sdr["UplData"];

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
        public IActionResult IPShipment()
        {
            //IPlabel temp = new IPlabel();
            //IPRange range = new IPRange();
            //range.EndIP = "";
            //range.StartIP = "";
            //List<IPRange> list = new List<IPRange>();
            //list.Add(range);
            //temp.Iprange = list;
            
            return View();
        }

        //[HttpPost]
        //public IActionResult IPShipment(IPlabel Request)
        //{
        //    Request.Iprange.Add(new IPRange());

        //    return View(Request);
        //}

        [HttpPost]
        public IActionResult IPShipment(IPlabel Request, string username, string password, string SiteID)
        {
            string userid = HttpContext.Session.GetString("suserid");
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
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
                        return View();
                    }
                    chkStr = startIPArray[i].Substring(1, 4);
                    chkEnd = endIPArray[i].Substring(1, 4);
                    if (!int.TryParse(chkStr, out int result) || !int.TryParse(chkEnd, out int result1))
                    {
                        TempData["ErrorMessage"] = "Kit number format is incorrect.";
                        return View();
                    }
                    else
                    {
                        intkStr = result;
                        intEnd = result1;
                    }
                   // string range = startIPArray[i] + "-" + endIPArray[i];
                    //if (seenRanges.Contains(range))
                    //{
                    //    TempData["ErrorMessage"] = "Part of the entered kit range is duplicate. Please verify and update.";
                    //    return View();
                    //}
                    //seenRanges.Add(range);
                    //if ((intkStr > 0) && (intEnd > 0))
                    //{
                    //    if (intkStr > intEnd)
                    //    {
                    //        TempData["ErrorMessage"] = "Start Kit Number cannot be greater than End.";
                    //        return View();
                    //    }

                    //    // if (i > 0 && intkStr <= previousEnd)
                    //    if (i > 0 && intkStr >= previousStart && intkStr <= previousEnd)
                    //    {
                    //        TempData["ErrorMessage"] = "Part of the entered kit range is duplicate. Please verify and update.";
                    //        return View();
                    //    }

                    //    previousStart = intkStr; // Update previous start value for the next iteration
                    //    previousEnd = intEnd;    // Update previous end value for the next iteration

                    //}


                    if (intkStr > 0 && intEnd > 0)
                   {
                        if (intkStr > intEnd)
                        {
                            TempData["ErrorMessage"] = "Start Kit Number cannot be greater than End.";
                            return View();
                        }

                        string range = startIPArray[i] + "-" + endIPArray[i]; // Construct range string
                            foreach (string seenRange in seenRanges)
                            {
                                if (CheckFullRangeOverlap(range, seenRange))
                                {
                                    TempData["ErrorMessage"] = "Part of the entered kit range is duplicate. Please verify and update.";
                                    return View();
                                }
                            }
                            seenRanges.Add(range);

                            if (CheckRangeOverlap(intkStr, intEnd, previousStart, previousEnd))
                            {
                                TempData["ErrorMessage"] = "Part of the entered kit range is duplicate. Please verify and update.";
                                return View();
                            }
                            previousStart = intkStr;
                            previousEnd = intEnd;
                    }
                        
                  


                    rtnVal = CheckBox(startIPArray[i], endIPArray[i]);
                    if (rtnVal == "")
                    {

                        TempData["ErrorMessage"] = "Part of the entered Kit range is not valid. Please verify and update.";
                        return View();
                    }

                    rtnVal = ChkLblToShipDepot(startIPArray[i], endIPArray[i]);
                    if (rtnVal != "")
                    {
                        //Message.Text += "<br />Error: IP already has ship request. <br />" + rtnVal;

                        //TempData["ErrorMessage"] = "Part of the IP/Kit Range: " + startIPArray[i] + " - " + endIPArray[i] + " have already been shipped. Please check";
                        TempData["ErrorMessage"] = "Part of the entered kit range have already been shipped to the depot. Please verify and update.";
                        return View();
                    }

                    //rtnVal = Chkkits(startIPArray[i], endIPArray[i]);
                    //if (rtnVal == "")
                    //{
                    //    //Message.Text += "<br />Error: IP already has ship request. <br />" + rtnVal;

                    //    //TempData["ErrorMessage"] = "Part of the IP/Kit Range: " + startIPArray[i] + " - " + endIPArray[i] + " have already been shipped. Please check";
                    //    TempData["ErrorMessage"] = "Part of the entered kit range do not exists. Please verify and update.";
                    //    return View();
                    //}


                }

                    var msgBody = "";
                var rtnValue = "";
                var message = "";
                int count = 0;
                for (int i = 0; i < startIPArray.Length; i++)
                {
                   rtnValue = UpdtLabelShip(startIPArray[i], endIPArray[i], Request.IPLblShipDTC, Request.Courier, Request.TrackNo, Request.IPLblShipExpiryDate, Request.IPLblShipLotNo, userid);
                    if((rtnValue == "IP Label Ship Error")|| (rtnValue == "IP_LABEL_SHIP_REQ INSERT") || (rtnValue == "Error INSERT"))
                    {
                        TempData["ErrorMessage"] = rtnValue;
                        return View();
                    }
                    else
                    {
                        ///select
                        count += RowsAffect(startIPArray[i], endIPArray[i]);
                        message += startIPArray[i] + " to " + endIPArray[i] ;
                        if (i < startIPArray.Length - 1)
                        {
                            message += ", ";
                        }

                    }


                }
                message = "A total of " + count + " kit(s) will be shipped to the depot. Range(s) from " + message +".";
                TempData["Message"] = message;
                string randemail = GetProfVpe(SPKEY, "PMIPLblRel") + ";" + GetUserEmail();
                string subject = StudyName() + " - WebView RTSM - Kit shipment confirmation";
                msgBody += Environment.NewLine + "Protocol: " + StudyName();
                msgBody += Environment.NewLine + "This email is to notify you that the following kit shipment is ready to be sent to the depot, details below." +"\n";
                msgBody += Environment.NewLine + "Date of IP Shipment: " + Request.IPLblShipDTC;
                msgBody += Environment.NewLine + "Courier: " + Request.Courier;
                msgBody += Environment.NewLine + "Tracking Number: " + Request.TrackNo;
                msgBody += Environment.NewLine + "Expiry Date: " + Request.IPLblShipExpiryDate;
                msgBody += Environment.NewLine + "Lot Number: " + Request.IPLblShipLotNo;
                msgBody += Environment.NewLine + message;
                SendEmail(randemail + ";" + "sidran@amarexcro.com", subject, msgBody);
                return RedirectToAction("IPLabellingHome");
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


        public string ChkLblToShipDepot(string strIP, string endIP)
        {
            var rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();
                    var sqlState = "";
                    //sqlState = "SELECT KitNumber FROM BIL_IP_RANGE  WHERE  (([IPLabel] is null) or ([ILSKEY] is not null)) ";
                    sqlState = "SELECT KitNumber FROM BIL_IP_RANGE  WHERE  (([ILSKEY] is not null)) ";
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
                                rtnVal += ";" + rdr["KitNumber"].ToString();
                            }
                        }     
                
            }
            conn.Close();
            return rtnVal;
        }

        //Chkkitst(startIPArray[i], endIPArray[i]);

        //public string Chkkits(string strIP, string endIP)
        //{
        //    var rtnVal = "";
        //    connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
        //    SqlConnection conn = new SqlConnection(connectionString);
        //    using (conn)
        //    {
        //        conn.Open();
        //        var sqlState = "";
        //        //sqlState = "SELECT KitNumber FROM BIL_IP_RANGE  WHERE  (([IPLabel] is null) or ([ILSKEY] is not null)) ";
        //        //sqlState = "SELECT KitNumber FROM BIL_IP_RANGE  WHERE  (([ILSKEY] is null)) ";
        //        //sqlState += "AND ([KitNumber] between '" + strIP + "' AND '" + endIP + "')";
        //        sqlState = "SELECT [KitNumber] FROM BIL_IP_RANGE WHERE [ILSKEY] IS NULL AND ([KitNumber] = '" + strIP + "' OR [KitNumber] = '" + endIP + "')";
        //        SqlCommand cmd = new SqlCommand(sqlState, conn);
        //        SqlDataReader rdr = cmd.ExecuteReader();
        //        while (rdr.Read())
        //        {
        //            if (rtnVal == "")
        //            {
        //                rtnVal = rdr["KitNumber"].ToString();
        //            }
        //            else
        //            {
        //                rtnVal += ";" + rdr["KitNumber"].ToString();
        //            }
        //        }

        //    }
        //    conn.Close();
        //    return rtnVal;
        //}
        public string UpdtLabelShip(string strIP, string endIP, string shipdtc, string courier, string trkno, string expiry, string lotNo, string uid)
        {
            string rtnVal = "IP Label Ship Error";
            var chkILSkey = "";
            chkILSkey = ChkKeyils(trkno, expiry, lotNo);
            var ilskey = "";
            if (chkILSkey == "NF")
            {
                ilskey = InsIPShipTrkExpiryLot(courier, trkno, expiry, lotNo, shipdtc, uid);
                if (ilskey == "Error INSERT")
                {
                    rtnVal = "Error INSERT";
                    return rtnVal;
                }
            }
            else
            {
                ilskey = chkILSkey;
            }
            rtnVal = InsIPShipReq(ilskey, courier, trkno, expiry, lotNo, strIP, endIP, uid);
            if (rtnVal == "IP_LABEL_SHIP_REQ INSERT")
            {
                return rtnVal;
            }
            string connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                
                    string sqlState = "";
                    
                    sqlState = "UPDATE BIL_IP_RANGE SET ILSKEY = @ILSKEY, [IPLblShipDTC] = @IPLblShipDTC, Courier = @Courier, TrackNo = @TrackNo, IPLblShipExpiryDate = @IPLblShipExpiryDate, IPLblShipLotNo = @IPLblShipLotNo, [IPLblShipUID] = @IPLblShipUID, [IPLblShipSysDTC] = sysdatetime() WHERE ([SITEID] is null) AND ([KitNumber] between '" + strIP + "' AND '" + endIP + "')";
                    SqlCommand cmd = new SqlCommand(sqlState, conn);
                    cmd.Parameters.AddWithValue("@ILSKEY", ilskey);
                    cmd.Parameters.AddWithValue("@IPLblShipDTC", shipdtc);
                    cmd.Parameters.AddWithValue("@Courier", courier);
                    cmd.Parameters.AddWithValue("@TrackNo", trkno);
                    cmd.Parameters.AddWithValue("@IPLblShipExpiryDate", expiry);
                    cmd.Parameters.AddWithValue("@IPLblShipLotNo", lotNo);
                    cmd.Parameters.AddWithValue("@IPLblShipUID", uid);

                     int rowsAfft = (int)cmd.ExecuteNonQuery();
                    rtnVal = "A total of " + rowsAfft + " kit(s) will be shipped to the depot.";
                
            }
            return rtnVal;
        }

        public  string ChkKeyils(string trkno, string expiry, string lotNo)
        {
            string retVal = "NF";

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();
                string sqlStat;
                sqlStat = "SELECT * FROM IP_LABEL_SHIP WHERE (TrackNo = '" + trkno + "') AND (ExpiryDate = '" + expiry + "') AND (LotNo = '" + lotNo + "')";
                SqlCommand cmd = new SqlCommand(sqlStat, conn);
                SqlDataReader rdr = cmd.ExecuteReader();
                    if (rdr.Read())
                    {
                        retVal = rdr["ILSKEY"].ToString();
                    }

            }
            conn.Close();
            return retVal;
        }

        public string InsIPShipReq(string ilskey, string courier, string trkno, string expiry, string lotNo, string rangestr, string rangeend, string uid)
        {
            string rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();
               
                    string sqlState = "";
                    sqlState = "INSERT INTO [IP_LABEL_SHIP_REQ] (ILSKEY, Courier, [TrackNo], [ExpiryDate], [LotNo], IPLabelShip, RangeStr, RangeEnd, [ADDUSER], ADDDATE) VALUES (@ILSKEY, @Courier, @TrackNo, @ExpiryDate, @LotNo, @IPLabelShip, @RangeStr, @RangeEnd, @ADDUSER, @ADDDATE)";
                    SqlCommand cmd = new SqlCommand(sqlState, conn);
                cmd.Parameters.AddWithValue("@ADDDATE", DateTime.Now);
                cmd.Parameters.AddWithValue("@ILSKEY", ilskey);
        
                cmd.Parameters.AddWithValue("@Courier", courier);
                cmd.Parameters.AddWithValue("@TrackNo", trkno);
                cmd.Parameters.AddWithValue("@ExpiryDate", expiry);
                cmd.Parameters.AddWithValue("@LotNo", lotNo);
                cmd.Parameters.AddWithValue("@IPLabelShip", "YES");
                cmd.Parameters.AddWithValue("@RangeStr", rangestr);
                cmd.Parameters.AddWithValue("@RangeEnd", rangeend);
                cmd.Parameters.AddWithValue("@ADDUSER", uid);
               
                int rowsAfft = (int)cmd.ExecuteNonQuery();
                    if (rowsAfft <= 0)
                    {
                        rtnVal = "IP_LABEL_SHIP_REQ INSERT";
                    }
                
            }
            return rtnVal;
        }

        public string InsIPShipTrkExpiryLot(string courier, string trkno, string expiry, string lotNo, string shipdtc, string uid)
        {
            string rtnVal = "";
            string connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = conn;
                    string sqlState = "INSERT INTO [IP_LABEL_SHIP] (Courier, [TrackNo], [ExpiryDate], [LotNo], IPLabelShip, IPLblShipDTC, [ADDUSER]) VALUES (@Courier, @TrackNo, @ExpiryDate, @LotNo, 'YES', @ShipDTC, @UID); SELECT @keyils = SCOPE_IDENTITY()";
                    cmd.CommandText = sqlState;

                    // Add parameters
                    cmd.Parameters.AddWithValue("@Courier", courier);
                    cmd.Parameters.AddWithValue("@TrackNo", trkno);
                    cmd.Parameters.AddWithValue("@ExpiryDate", expiry);
                    cmd.Parameters.AddWithValue("@LotNo", lotNo);
                    cmd.Parameters.AddWithValue("@ShipDTC", shipdtc);
                    cmd.Parameters.AddWithValue("@UID", uid);

                    // Output parameter
                    SqlParameter pRV = new SqlParameter("@keyils", SqlDbType.Int);
                    pRV.Direction = ParameterDirection.Output;
                    cmd.Parameters.Add(pRV);

                    // Execute query
                    int rowsAfft = cmd.ExecuteNonQuery();

                    if (rowsAfft <= 0)
                    {
                        rtnVal = "Error INSERT";
                    }
                    else
                    {
                        rtnVal = cmd.Parameters["@keyils"].Value.ToString();
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
        public IActionResult EditShipment(int ILSKEY)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string selectSql = "SELECT Courier, TrackNo, ExpiryDate, LotNo, IPLblShipDTC, ILSKEY, ADDUSER FROM IP_LABEL_SHIP WHERE ILSKEY = " + ILSKEY +"";
            IPShip temp = new IPShip();
            temp.ILSKEY = ILSKEY;

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // Retrieve existing data
                using (SqlCommand selectCmd = new SqlCommand(selectSql, con))
                {
                    SqlDataReader rdr = selectCmd.ExecuteReader();
                    if (rdr.Read())
                    {
                        temp.Courier = rdr["Courier"].ToString();
                        temp.TrackNo = rdr["TrackNo"].ToString();
                        temp.ExpiryDate = rdr["ExpiryDate"].ToString();
                        temp.LotNo = rdr["LotNo"].ToString();
                        temp.IPLblShipDTC = rdr["IPLblShipDTC"].ToString();
                        temp.ILSKEY = (int)rdr["ILSKEY"];
                        temp.ADDUSER = rdr["ADDUSER"].ToString();
                        temp.IPDetails = GetIPDetails(ILSKEY);
                    }
                    rdr.Close();
                }
                con.Close();
            }
            return View(temp);
        }

        [HttpPost]
        public IActionResult EditShipment(int ILSKEY, string IPLblShipDTC, string Courier, string TrackNo, string ExpiryDate, string LotNo, string username, string password)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string selectSql = "SELECT Courier, TrackNo, ExpiryDate, LotNo, IPLblShipDTC, ILSKEY, ADDUSER FROM IP_LABEL_SHIP WHERE ILSKEY = " + ILSKEY + "";
            IPShip temp = new IPShip();
            temp.ILSKEY = ILSKEY;

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // Retrieve existing data
                using (SqlCommand selectCmd = new SqlCommand(selectSql, con))
                {
                    SqlDataReader rdr = selectCmd.ExecuteReader();
                    if (rdr.Read())
                    {
                        temp.Courier = rdr["Courier"].ToString();
                        temp.TrackNo = rdr["TrackNo"].ToString();
                        temp.ExpiryDate = rdr["ExpiryDate"].ToString();
                        temp.LotNo = rdr["LotNo"].ToString();
                        temp.IPLblShipDTC = rdr["IPLblShipDTC"].ToString();
                        temp.ILSKEY = (int)rdr["ILSKEY"];
                        temp.ADDUSER = rdr["ADDUSER"].ToString();
                        temp.IPDetails = GetIPDetails(ILSKEY);
                    }
                    rdr.Close();
                }
                con.Close();
            }

            if (IsValidUser(username, password) && (string.Equals(userid, username, StringComparison.OrdinalIgnoreCase) || userid == "sidran"))
            {   
                string rtnval = ChkAtDepot(ILSKEY);
                if(rtnval != "OK")
                {
                    TempData["ErrorMessage"] = rtnval;
                    return View(temp);

                }
                string val = UpdtIPLabel(ILSKEY, IPLblShipDTC, Courier, TrackNo, ExpiryDate, LotNo, userid);
                if(val != "OK")
                {
                    TempData["ErrorMessage"] = "IP Label Update Error";
                    return View(temp);
                }
                TempData["Message"] = "Shipment information has been updated successfully.";
                string randemail = GetProfVpe(SPKEY, "PMIPLblRel") + ";" + GetUserEmail();
                string subject = StudyName() + " - WebView RTSM - Kit shipment update confirmation";
                string msgBody = "";
                msgBody += Environment.NewLine + "Protocol: " + StudyName();
                msgBody += Environment.NewLine + "This email is to notify you that the following kit(s) shipment has been updated, details below." + "\n";
                msgBody += Environment.NewLine + "Lot Number: " + LotNo ;
                msgBody += Environment.NewLine + "Tracking Number: " + TrackNo;
                msgBody += Environment.NewLine + "Shipment Date: " + IPLblShipDTC;
                msgBody += Environment.NewLine + "Expiry date: " + ExpiryDate;

                SendEmail(randemail + ";" + "sidran@amarexcro.com" , subject, msgBody);
                return RedirectToAction("IPLabellingHome");
                
            }
            else
            {
                string errorMessage = "Invalid username or password.";
                ViewBag.ErrorMessage = errorMessage;
                TempData["ErrorMessage"] = errorMessage;

                return View(temp);
            }
                
        }


        //Edit Expiry Date only

        public IActionResult EditExpiry(int ILSKEY)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string selectSql = "SELECT Courier, TrackNo, ExpiryDate, LotNo, IPLblShipDTC, ILSKEY, ADDUSER FROM IP_LABEL_SHIP WHERE ILSKEY = " + ILSKEY + "";
   
            IPShip temp = new IPShip();
            temp.ILSKEY = ILSKEY;

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // Retrieve existing data
                using (SqlCommand selectCmd = new SqlCommand(selectSql, con))
                {
                    SqlDataReader rdr = selectCmd.ExecuteReader();
                    if (rdr.Read())
                    {
                        temp.Courier = rdr["Courier"].ToString();
                        temp.TrackNo = rdr["TrackNo"].ToString();
                        temp.ExpiryDate = rdr["ExpiryDate"].ToString();
                        temp.LotNo = rdr["LotNo"].ToString();
                        temp.IPLblShipDTC = rdr["IPLblShipDTC"].ToString();
                        temp.ILSKEY = (int)rdr["ILSKEY"];
                        temp.ADDUSER = rdr["ADDUSER"].ToString();
                        temp.IPDetails = GetIPDetails(ILSKEY);

                        //temp.SFDATE = rdr["SFDATE"].ToString();
                    }
                    rdr.Close();
                }
                con.Close();
            }
            return View(temp);
        }

        [HttpPost]
        public IActionResult EditExpiry(int ILSKEY, string IPLblShipDTC, string Courier, string TrackNo, string ExpiryDate, string LotNo, string username, string password)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string selectSql = "SELECT Courier, TrackNo, ExpiryDate, LotNo, IPLblShipDTC, ILSKEY, ADDUSER FROM IP_LABEL_SHIP WHERE ILSKEY = " + ILSKEY + "";
            IPShip temp = new IPShip();
            temp.ILSKEY = ILSKEY;

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // Retrieve existing data
                using (SqlCommand selectCmd = new SqlCommand(selectSql, con))
                {
                    SqlDataReader rdr = selectCmd.ExecuteReader();
                    if (rdr.Read())
                    {
                        temp.Courier = rdr["Courier"].ToString();
                        temp.TrackNo = rdr["TrackNo"].ToString();
                        temp.ExpiryDate = rdr["ExpiryDate"].ToString();
                        temp.LotNo = rdr["LotNo"].ToString();
                        temp.IPLblShipDTC = rdr["IPLblShipDTC"].ToString();
                        temp.ILSKEY = (int)rdr["ILSKEY"];
                        temp.ADDUSER = rdr["ADDUSER"].ToString();
                        temp.IPDetails = GetIPDetails(ILSKEY);
                    }
                    rdr.Close();
                }
                con.Close();
            }
            if (IsValidUser(username, password) && (string.Equals(userid, username, StringComparison.OrdinalIgnoreCase) || userid == "sidran"))
            {
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                //string rtnval = ChkAtDepot(ILSKEY);
                //if (rtnval != "OK")
                //{
                //    TempData["ErrorMessage"] = rtnval;
                //    return View();

                //}
                string val = UpdtExpiryDate(ILSKEY, ExpiryDate, userid);
                if (val != "OK")
                {
                    TempData["ErrorMessage"] = "IP Label Update Error";
                    return View(temp);
                }
                TempData["Message"] = "Expirty date has been updated successfully.";
                string randemail = GetProfVpe(SPKEY, "PMIPLblRel") + ";" + GetUserEmail();
                string subject = StudyName() + " - WebView RTSM - Kit expiry date update confirmation";
                string msgBody = "";
                msgBody += Environment.NewLine + "Protocol: " + StudyName() + "\n";
                msgBody +=  "This email is to notify you that the following kit(s) expiry date has been updated, details below.";
                msgBody +=  "Lot Number: " + LotNo + "\n";
                msgBody += "Tracking Number: " + TrackNo + "\n";
                msgBody +=  "Shipment Date: " + IPLblShipDTC + "\n";
                msgBody += "Expiry date: " + ExpiryDate + "\n";

                SendEmail(randemail + ";" + "sidran@amarexcro.com", subject, msgBody);
                return RedirectToAction("IPLabellingHome");

            }
            else
            {
                string errorMessage = "Invalid username or password.";
                ViewBag.ErrorMessage = errorMessage;
                TempData["ErrorMessage"] = errorMessage;

                return View(temp);
            }

        }


        public string ChkAtDepot(int ilskey)
        {
            string retVal = "OK";

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            string sqlState = "SELECT COUNT(*) FROM [BIL_IP_RANGE] WHERE (ILSKEY = " + ilskey + ") AND (([ATDEPOT] is not null))";
            SqlCommand cmd = new SqlCommand(sqlState, conn);
            using (conn)
            {
                conn.Open();
                int count = (int)cmd.ExecuteScalar();
                if (count > 0)
                {
                  retVal = "Can not Modify or Cancel Shipment Processed at Depot ";
                }
                
            }
            conn.Close();
            return retVal;
        }

        public string UpdtIPLabel(int ilskey, string shipdtc, string courier, string trkno, string expiry, string lotNo, string uid)
        {
            string rtnVal = "Error: IP Label Update Error";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            string sqlState = "UPDATE BIL_IP_RANGE SET [IPLblShipDTC] = @IPLblShipDTC, [Courier] = @Courier, [TrackNo] = @TrackNo, [IPLblShipExpiryDate] = @IPLblShipExpiryDate, [IPLblShipLotNo] = @IPLblShipLotNo, [IPLblShipChangeUID] = @IPLblShipChangeUID, IPLblShipChangeDTC = @IPLblShipChangeDTC WHERE ([ILSKEY] = " + ilskey + ")";
            string sqlState2 = "UPDATE [IP_LABEL_SHIP] SET [IPLblShipDTC] = @IPLblShipDTC, [Courier] = @Courier, [TrackNo] = @TrackNo, [ExpiryDate] = @ExpiryDate, [LotNo] = @LotNo, [CHANGEUSER] = @CHANGEUSER, [CHANGEDATE] = @CHANGEDATE WHERE ([ILSKEY] = " + ilskey + ")";
            string sqlState4 = "UPDATE [IP_LABEL_SHIP_REQ] SET  [ExpiryDate] = @ExpiryDate, [Courier] = @Courier, [TrackNo] = @TrackNo, [LotNo] = @LotNo, [CHANGEUSER] = @CHANGEUSER, [CHANGEDATE] = @CHANGEDATE WHERE ([ILSKEY] = " + ilskey + ")";
            SqlCommand cmd = new SqlCommand(sqlState, conn);
            SqlCommand cmd2 = new SqlCommand(sqlState2, conn);
            SqlCommand cmd4 = new SqlCommand(sqlState4, conn);
            conn.Open();
            cmd.Parameters.AddWithValue("@IPLblShipDTC", shipdtc);
            cmd.Parameters.AddWithValue("@Courier", courier);
            cmd.Parameters.AddWithValue("@TrackNo", trkno);
            cmd.Parameters.AddWithValue("@IPLblShipExpiryDate", expiry);
            cmd.Parameters.AddWithValue("@IPLblShipLotNo", lotNo);
            cmd.Parameters.AddWithValue("@IPLblShipChangeUID", uid);
            cmd.Parameters.AddWithValue("@IPLblShipChangeDTC", DateTime.Now);
            cmd.ExecuteNonQuery();

            cmd2.Parameters.AddWithValue("@IPLblShipDTC", shipdtc);
            cmd2.Parameters.AddWithValue("@Courier", courier);
            cmd2.Parameters.AddWithValue("@TrackNo", trkno);
            cmd2.Parameters.AddWithValue("@ExpiryDate", expiry);
            cmd2.Parameters.AddWithValue("@LotNo", lotNo);
            cmd2.Parameters.AddWithValue("@CHANGEUSER", uid);
            cmd2.Parameters.AddWithValue("@CHANGEDATE", DateTime.Now);
            cmd2.ExecuteNonQuery();


            
            cmd4.Parameters.AddWithValue("@Courier", courier);
            cmd4.Parameters.AddWithValue("@TrackNo", trkno);
            cmd4.Parameters.AddWithValue("@ExpiryDate", expiry);
            cmd4.Parameters.AddWithValue("@LotNo", lotNo);
            cmd4.Parameters.AddWithValue("@CHANGEUSER", uid);
            cmd4.Parameters.AddWithValue("@CHANGEDATE", DateTime.Now);
            cmd4.ExecuteNonQuery();
            rtnVal = "OK";
            conn.Close();
            return rtnVal;
        }


        public string UpdtExpiryDate(int ilskey, string expiry, string uid)
        {
            string rtnVal = "Error: IP Label Update Error";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            string sqlState = "UPDATE BIL_IP_RANGE SET  [IPLblShipExpiryDate] = '" + expiry + "', [IPLblShipChangeUID] = '" + uid + "', IPLblShipChangeDTC = sysdatetime() WHERE ([ILSKEY] = " + ilskey + ")";
            string sqlState2 = "UPDATE [IP_LABEL_SHIP] SET [ExpiryDate] = '" + expiry + "', [CHANGEUSER] = '" + uid + "', [CHANGEDATE] = sysdatetime() WHERE ([ILSKEY] = " + ilskey + ")";
            string sqlState3 = "UPDATE IP_LABEL_DEPOT_REL SET  [ExpiryDate] = '" + expiry + "', [CHANGEUSER] = '" + uid + "', [CHANGEDATE] = sysdatetime() WHERE ([ILSKEY] = " + ilskey + ")";
            string sqlState4 = "UPDATE [IP_LABEL_SHIP_REQ] SET [ExpiryDate] = '" + expiry + "', [CHANGEUSER] = '" + uid + "', [CHANGEDATE] = sysdatetime() WHERE ([ILSKEY] = " + ilskey + ")";
            SqlCommand cmd = new SqlCommand(sqlState, conn);
            SqlCommand cmd2 = new SqlCommand(sqlState2, conn);
            SqlCommand cmd3 = new SqlCommand(sqlState3, conn);
            SqlCommand cmd4 = new SqlCommand(sqlState4, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            cmd2.ExecuteNonQuery();
            cmd3.ExecuteNonQuery();
            cmd4.ExecuteNonQuery();
            rtnVal = "OK";
            conn.Close();
            return rtnVal;
        }

        public IActionResult CancelShipment(int ILSKEY, string LotNo, string TrackNo, string IPLblShipDTC, string username, string password, string ExpiryDate)
        {
            string userid = HttpContext.Session.GetString("suserid");
            if (IsValidUser(username, password) && (string.Equals(userid, username, StringComparison.OrdinalIgnoreCase) || userid == "sidran"))
            {
                int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
                
                connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                string rtnval = ChkAtDepot(ILSKEY);
                if (rtnval != "OK")
                {
                    TempData["ErrorMessage"] = rtnval;
                    return RedirectToAction("IPLabellingHome");

                }
                string val = UpdtCanShip(ILSKEY, userid);
                if (val != "OK")
                {
                    TempData["ErrorMessage"] = "Error occured in cancelling Shipment";
                    return RedirectToAction("IPLabellingHome");
                }
                TempData["Message"] = "Shipment has been canceled successfully.";
                string randemail = GetProfVpe(SPKEY, "PMIPLblRel") + ";" + GetUserEmail();
                string subject = "Webview RTSM - IP Label Shipment Canceled For [Amarex][Webview RTSM - Test]";
                string msgBody = "";
                msgBody += Environment.NewLine + "Protocol: WebView RTSM-Test";
                msgBody += Environment.NewLine + "This email is to notify you that the following IP shipment has been canceled" + Environment.NewLine; ;
                msgBody += Environment.NewLine + "Lot Number: " + LotNo + Environment.NewLine; 
                msgBody += Environment.NewLine + "Tracking Number: " + TrackNo + Environment.NewLine;
                msgBody += Environment.NewLine + "Shipment Date: " + IPLblShipDTC + Environment.NewLine;
                msgBody += Environment.NewLine + "Expiry date: " + ExpiryDate + Environment.NewLine; 


                SendEmail(randemail + ";" + "sidran@amarexcro.com", subject, msgBody);
                return RedirectToAction("IPLabellingHome");

            }
            else
            {
                string errorMessage = "Invalid username or password.";
                ViewBag.ErrorMessage = errorMessage;
                TempData["ErrorMessage"] = errorMessage;

                return RedirectToAction("IPLabellingHome");
            }

        }

        public string UpdtCanShip(int ilskey, string uid)
        {
            string rtnVal = "IP Label Remove Error";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                
                    string sqlState = "UPDATE BIL_IP_RANGE SET [ILSKEY] = null, [IPLblShipDTC] = null, [Courier] = null, [TrackNo] = null, [IPLblShipExpiryDate] = null, [IPLblShipLotNo] = null, [IPLblShipSysDTC] = null, [IPLblShipUID] = null  WHERE ([ILSKEY] = " + ilskey + "); ";
                    string sqlState2 = "UPDATE IP_LABEL_SHIP SET IP_LABEL_SHIP.CHANGEUSER = '" + uid + "', [CHANGEDATE] = sysdatetime(), [IPSTAT] = 'Cancel' WHERE ([ILSKEY] = " + ilskey + ")";
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

         public IActionResult ViewShipment(int ILSKEY)
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string selectSql = "SELECT Courier, TrackNo, ExpiryDate, LotNo, IPLblShipDTC, ILSKEY, ADDUSER FROM IP_LABEL_SHIP WHERE ILSKEY = " + ILSKEY +"";
            string sql = "SELECT RangeStr, RangeEnd, TELKEY FROM IP_LABEL_SHIP_REQ WHERE ILSKEY = " + ILSKEY +" ORDER BY RangeStr, TELKEY";

            IPShip temp = new IPShip();
            temp.ILSKEY = ILSKEY;
            var Req = new List<IPShipReq>();

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();

                // Retrieve existing data
                    SqlCommand selectCmd = new SqlCommand(selectSql, con);
                    SqlCommand cmd = new SqlCommand(sql, con);
                   SqlDataReader rdr = selectCmd.ExecuteReader();
                
                    while (rdr.Read())
                    {
                        temp.Courier = rdr["Courier"].ToString();
                        temp.TrackNo = rdr["TrackNo"].ToString();
                        temp.ExpiryDate = rdr["ExpiryDate"].ToString();
                        temp.LotNo = rdr["LotNo"].ToString();
                        temp.IPLblShipDTC = rdr["IPLblShipDTC"].ToString();
                        temp.ILSKEY = (int)rdr["ILSKEY"];
                        temp.ADDUSER = rdr["ADDUSER"].ToString();
                    }
                    rdr.Close();
                SqlDataReader rdr2 = cmd.ExecuteReader();
                while (rdr2.Read())
                    {
                        var temp1 = new IPShipReq();
                        temp1.RangeStr = rdr2["RangeStr"].ToString();
                        temp1.TELKEY = (int)rdr2["TELKEY"];
                        temp1.RangeEnd = rdr2["RangeEnd"].ToString();
                    Req.Add(temp1);
                    }
                    rdr2.Close();

                con.Close();
            }
            temp.IPReq = Req;
            return View(temp);
        }

       // report.fileList = GetFiles(key);

        public List<ShipDetails> GetIPDetails(int ILSKEY) 
        {
            // ViewBag.studyKey = StudyKey;
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
                    SELECT ', ' + CONCAT(depot.RangeStr, ' - ', depot.RangeEnd, ' $ ', depot.IPDepotDTC, ' $ ', depot.ADDUSER,'')
                    FROM IP_LABEL_DEPOT_REL AS depot
                    WHERE ship.ILSKEY = depot.ILSKEY
                    FOR XML PATH('')), 1, 2, '') AS DepotRangeStrEnd
            FROM 
                IP_LABEL_SHIP AS ship
            LEFT JOIN 
                IP_LABEL_DEPOT_REL AS depot ON ship.ILSKEY = depot.ILSKEY
            WHERE ship.ILSKEY = @ILSKEY
                
            GROUP BY 
                ship.Courier, 
                ship.TrackNo, 
                ship.ExpiryDate, 
                ship.LotNo, 
                ship.IPLblShipDTC, 
                ship.ILSKEY, 
                ship.ADDUSER
            ORDER BY 
                ship.ILSKEY;";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            ShipDetails IPdetails = new ShipDetails();
            var model = new List<ShipDetails>();

            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@ILSKEY", ILSKEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var temp = new ShipDetails();
                    temp.ILSKEY = (int)rdr["ILSKEY"];
                    temp.TrackNo = rdr["TrackNo"].ToString();
                    temp.Courier = rdr["Courier"].ToString();
                    temp.ExpiryDate = rdr["ExpiryDate"].ToString();
                    temp.LotNo = rdr["LotNo"].ToString();
                    temp.ADDUSER = rdr["ADDUSER"].ToString();
                    temp.IPLblShipDTC = rdr["IPLblShipDTC"].ToString();
                    temp.RangeStrEnd = rdr["RangeStrEnd"].ToString();
                    temp.DepotRangeStrEnd = rdr["DepotRangeStrEnd"].ToString(); 
                    model.Add(temp);
                }
            }
            conn.Close();
            return model; 
        }

        public int RowsAffect(string startIP, string endIP)
        {
            int count = 0;
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            using (conn)
            {
                conn.Open();

                string sqlState = "SELECT COUNT(*) FROM BIL_IP_RANGE WHERE IPLblShipDTC is not null AND ([KitNumber] BETWEEN '" + startIP + "' AND '" + endIP + "')";
                SqlCommand cmd = new SqlCommand(sqlState, conn);
                count = (int)cmd.ExecuteScalar();

            }
            conn.Close();
            return count;
        }

    }
}
