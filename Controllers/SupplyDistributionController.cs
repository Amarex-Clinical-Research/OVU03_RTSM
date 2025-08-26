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
    public class SupplyDistributionController : Controller
    {

        private readonly ILogger<SupplyDistributionController> _logger;
        public string connectionString;
        readonly IConfiguration _configuration;

        public SupplyDistributionController(ILogger<SupplyDistributionController> logger, IConfiguration configuration)
        {
            _logger = logger;
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
        public IActionResult SupplyDistributionHome()
        {

            var list = new SPModel();
            var site = new List<SiteID>();
            var sitekit = new List<ShipIP>();
            var depotkit = new List<IPStatus>();

            string sql = "SELECT DISTINCT SITEID, SITENAME, INVNAME FROM ShipToSite WHERE MISC1 is null AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID";
            string kitsql = "SELECT SITEID, KitNumber, SENDBY, REPLACE(UPPER(CONVERT(varchar, [SENDDATE], 106)), ' ', '/') AS 'REQ DATE', RECVDBY, REPLACE(UPPER(CONVERT(varchar, [RECVDDATE], 106)), ' ', '/') AS 'RECVD DATE', KITSTAT, ASSIGNED, ASSIGNMENT_DATE, SUBJID, KITCOMM, KitRepled, IPLblShipExpiryDate , IPLblShipLotNo ,KITSET, KITKEY FROM BIL_IP_RANGE WHERE (SITEID IS NOT NULL) ORDER BY SITEID, KitNumber, KITKEY";
            string depotsql = "SELECT KitNumber, KITSTAT, KITKEY, IPLblShipExpiryDate , IPLblShipLotNo , ATDEPOT, ATDEPOTDTC FROM BIL_IP_RANGE WHERE (SITEID IS NULL) AND (ATDEPOT = 'Yes') ORDER BY KitNumber, KITKEY";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            SqlCommand cmd1 = new SqlCommand(kitsql, conn);
            SqlCommand cmd2 = new SqlCommand(depotsql, conn);

            using (conn)
            {
                conn.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var siteinfo = new SiteID();
                    // siteinfo.key = (int)rdr["STSKEY"];
                    siteinfo.SITEID = rdr["SITEID"].ToString() + "-" + rdr["SITENAME"].ToString();
                    siteinfo.INVNAME = rdr["INVNAME"].ToString();


                    site.Add(siteinfo);

                }
                rdr.Close();
                SqlDataReader rdr1 = cmd1.ExecuteReader();
                while (rdr1.Read())
                {
                    var temp = new ShipIP();
                    temp.KITKEY = (int)rdr1["KITKEY"];
                    temp.SITEID = rdr1["SITEID"].ToString();
                    temp.KitNumber = rdr1["KitNumber"].ToString();
                    temp.KITSTAT = rdr1["KITSTAT"].ToString();
                    temp.RECVDBY = rdr1["RECVDBY"].ToString();
                    temp.RECVDDATEstr = rdr1["RECVD DATE"].ToString();
                    temp.SUBJID = rdr1["SUBJID"].ToString();
                    temp.SENDDATEstr = rdr1["REQ DATE"].ToString();
                    temp.SENDBY = rdr1["SENDBY"].ToString();
                    temp.ASSIGNED = rdr1["ASSIGNED"].ToString();
                    temp.ASSIGNMENT_DATE = rdr1["ASSIGNMENT_DATE"] as DateTime?;
                    temp.KITCOMM = rdr1["KITCOMM"].ToString();
                    temp.KitRepled= rdr1["KitRepled"].ToString();
                    temp.IPLblShipExpiryDate= rdr1["IPLblShipExpiryDate"].ToString();
                    temp.IPLblShipLotNo= rdr1["IPLblShipLotNo"].ToString();
                    if(rdr1["KITSET"] == DBNull.Value)
                    {
                         temp.KITSET = 0;
                    }
                    else
                        temp.KITSET = Convert.ToInt16(rdr1["KITSET"]);
                    //temp.KITSET = (short)rdr1["KITSET"];


                    sitekit.Add(temp);
                   
                }
                rdr1.Close();
                SqlDataReader rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read())
                {
                    var temp = new IPStatus();
                    temp.KitNumber = rdr2["KitNumber"].ToString();
                    temp.KITSTAT = rdr2["KITSTAT"].ToString();
                    temp.KITKEY = (int)rdr2["KITKEY"];
                    temp.IPLblShipExpiryDate = rdr2["IPLblShipExpiryDate"].ToString();
                    temp.IPLblShipLotNo = rdr2["IPLblShipLotNo"].ToString();
                    temp.ATDEPOT = rdr2["ATDEPOT"].ToString();
                    temp.ATDEPOTDTC = rdr2["ATDEPOTDTC"] as DateTime?;

                    depotkit.Add(temp);

                }
                rdr2.Close();

            }
            conn.Close();
            list.Site = site;
            list.SiteKits = sitekit;
            list.DepotKits = depotkit;

            return View(list);
        }

        public IActionResult Update(SPModel Request, string SITEID, string shipNumber, string shiptype, string INVNAME)
        {

            if (!int.TryParse(shipNumber, out int ShipNumber))
            {
                TempData["ErrorMessage"] = "Number of shipments must be a valid number.";
                return RedirectToAction("SupplyDistributionHome");
            }

            int numBottles;
            string userid = HttpContext.Session.GetString("suserid");
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string[] arr = SITEID.Split("-");
            SITEID = arr[0];
            int count = VerifySiteInfo(SITEID, INVNAME, SPKEY);
            if(count != 1)
            {
                TempData["ErrorMessage"] = "Site ID and Investigator name does not match";
                return RedirectToAction("SupplyDistributionHome");

            }
            //if (shiptype == "Initial")
            //    numBottles = 16;
            //else
                numBottles = ShipNumber;
            DateTime dteNow = new DateTime();
            dteNow = DateTime.Now;
            string bottles = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sqlState = "DECLARE @KS int; SELECT @KS = MAX(KITSET) FROM BIL_IP_RANGE; IF @KS is null SELECT @KS = 1 ELSE SELECT @KS = @KS + 1; ";
            sqlState += "UPDATE BIL_IP_RANGE SET SITEID = '" + SITEID + "', KITSET = @KS, SENDBY = '" + userid + "', SENDDATE = '" + dteNow.ToString() + "', KITTYPE = '" + shiptype + "', KITSTAT = 'Shipped' WHERE KITKEY IN ";
            sqlState += "(SELECT TOP " + numBottles + " KITKEY FROM BIL_IP_RANGE WHERE ([TreatmentGroup] = 'ARM A') AND (SITEID IS NULL) AND ([ATDEPOT] is not NULL)  ORDER BY KitNumber); ";
            sqlState += "UPDATE BIL_IP_RANGE SET SITEID = '" + SITEID + "', KITSET = @KS, SENDBY = '" + userid + "', SENDDATE = '" + dteNow.ToString() + "', KITTYPE = '" + shiptype + "', KITSTAT = 'Shipped' WHERE KITKEY IN ";
            sqlState += "(SELECT TOP " + numBottles + " KITKEY FROM BIL_IP_RANGE WHERE ([TreatmentGroup] = 'ARM B') AND (SITEID IS NULL) AND ([ATDEPOT] is not NULL) ORDER BY KitNumber); ";
            string selectSql = "SELECT * FROM BIL_IP_RANGE WHERE (SITEID = '" + SITEID + "') AND (SENDBY = '" + userid + "') AND (SENDDATE = '" + dteNow.ToString() + "') AND (KITTYPE = '" + shiptype + "') AND (KITSTAT = 'Shipped')";
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sqlState, conn);
            SqlCommand cmd2 = new SqlCommand(selectSql, conn);
            var cntKits = 0;
            using (conn)
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                SqlDataReader rdr = cmd2.ExecuteReader();
                while (rdr.Read())
                {
                    bottles += rdr["KitNumber"].ToString() + "\n";
                    cntKits++;
                }
                rdr.Close();

            }
            if(cntKits <= 0)
            {
                TempData["ErrorMessage"] = "No Kit found";
                return RedirectToAction("SupplyDistributionHome");
            }
            DateTime dtmDate = DateTime.Now;
            string retSite = "";
            string retSupp = "";
            retSite = GetEmailByGrp(SPKEY, SITEID, "S");
            retSupp = GetEmailByGrp(SPKEY, "(All)", "D");
            //shipperEmail = BllDalGen.GetProfVpe(spkey, "ShipperEmail");

            string SendTo = GetEmailList(SPKEY, "PMIPLblRel") + ";" + GetEmailList(SPKEY, "IPDepot") + ";" + "jacobk@amarexcro.com" ;
            if(retSite != null)
            {
                SendTo = SendTo + ";" + retSite;
            }
            if(retSupp != null)
            {
                SendTo = SendTo + ";" + retSupp;
            }
            if (SendTo != null)
            {
                string subject = StudyName() + " - Webview RTSM - Request to Release Kit(s) for Site " + SITEID;
                string message = "Protocol: " + StudyName() + "\n";
                message += "Request to Release - package and ship" + "\n" + "\n";
                message += dtmDate.ToString("D") + "\n" + "\n";
                message += GetShipTo(SITEID, SPKEY) + "\n";
                message += "Total number of Kits Shipped: " + cntKits + "\n";
                message += "KitNumber in shipment:" + "\n" + bottles + "\n" + "\n";



                //retSite = GetEmailByGrp(spkey, siteid, "S");
                //retSupp = GetEmailByGrp(spkey, "(All)", "D");

                SendEmail(SendTo, subject, message);
                string emailStatus = " Email sent successfully!";
                TempData["Message"] = cntKits + " Kits shipped for Site ID :" + SITEID + "\n"+"\n" + emailStatus;
            }
            else
                TempData["Message"] = cntKits + " Kits shipped for Site ID :" + SITEID + " But No emails avaliable to notify of Shipment.";

            //TempData["ErrorMessage"] = "No kits found";
            return RedirectToAction("SupplyDistributionHome");
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

        public string GetEmailList(int SPKEY, string Type)
        {
            string sql = "Select PIDet from Email_Notifications where( SPKEY = " + SPKEY + " AND PIType = '" + Type + "')";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {

                    string result = cmd.ExecuteScalar() as string;

                    return result;
                }
            }


        }

        public string GetShipTo(string SiteID, int SPKEY)
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

        public IActionResult SiteKits(object sender, EventArgs e)

        {

            XLWorkbook wb = new XLWorkbook();
            DataTable dt = CreateDataTable("SELECT SITEID, KitNumber, SENDBY, REPLACE(UPPER(CONVERT(varchar, [SENDDATE], 106)), ' ', '/') AS 'REQ DATE', RECVDBY, REPLACE(UPPER(CONVERT(varchar, [RECVDDATE], 106)), ' ', '/') AS 'RECVD DATE', KITSTAT, ASSIGNED, ASSIGNMENT_DATE, SUBJID, KITCOMM, KitRepled, IPLblShipExpiryDate AS 'ExpiryDate', IPLblShipLotNo AS 'LotNo',KITSET, KITKEY FROM BIL_IP_RANGE WHERE (SITEID IS NOT NULL) ORDER BY SITEID, KitNumber, KITKEY");
            wb.Worksheets.Add(dt, "SiteKits").Columns().AdjustToContents(); // easiest way to convert sql data to a excel doc

            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "SiteKits" + DateTime.Now.Year + ".xlsx");

            }
        }

        public IActionResult DepotKits(object sender, EventArgs e)

        {

            XLWorkbook wb = new XLWorkbook();
            DataTable dt = CreateDataTable("SELECT KitNumber as 'Kit Number', IPLblShipExpiryDate AS 'Expiry Date', IPLblShipLotNo AS 'Lot No', ATDEPOT as 'At Depot', ATDEPOTDTC as 'AtDepot Date' FROM BIL_IP_RANGE WHERE (SITEID IS NULL) AND (ATDEPOT = 'Yes') ORDER BY KitNumber, KITKEY");
            wb.Worksheets.Add(dt, "DepotKits").Columns().AdjustToContents(); // easiest way to convert sql data to a excel doc

            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "DepotKits" + DateTime.Now.Year + ".xlsx");

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

        public int VerifySiteInfo(string SiteID, string Invname, int SPKEY)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            int count = 0;
            string sql = "SELECT COUNT(*) FROM ShipToSite WHERE (SITEID = '" + SiteID + "') AND (INVNAME = '" + Invname + "') AND (SPKEY = " + SPKEY + ")";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            using (con)
            {
                con.Open();
                count = (int)cmd.ExecuteScalar();


            }
            con.Close();

            return count;
        }

       
    }

}
