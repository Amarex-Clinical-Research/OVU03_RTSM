using AppUidAuth;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class RandSubj1
    {
        public int ROW_KEY { get; set; }
        public int SPKEY { get; set; }
        public string STATUS_INFO { get; set; }
        [Required(ErrorMessage = "This is a required field")]
        public string SITEID { get; set; }
        //[StringLength(3, MinimumLength = 3)]
        [Required]
        public string SUBJID { get; set; }
        //[RegularExpression(@"\d{4}", ErrorMessage = "Date format is YYYY")]
        //[Required(ErrorMessage = "Year of Birth is a required field")]
        public string BRTHDTC { get; set; }
        [BindProperty]
        //[Required(ErrorMessage = "Sex is a required field")]
        public string SEX { get; set; }
        [Required]
        public DateTime ICDTC { get; set; }
        public string ICDTCstr { get; set; }
        [Required]
        public DateTime SCRNDTC { get; set; }
        public string SCRNDTCstr { get; set; }
        //[Required(ErrorMessage = "Age Group is a required field")]
        public string AgeGroup { get; set; }
        //[Required(ErrorMessage = "Eligibility is a required field")]
        public string ELIGRAND { get; set; }
        public string StratumCode { get; set; }
        public string VISITID { get; set; }

        public string RANDBY { get; set; }

        public string DATERAND { get; set; }
        public string DateSponsorApproved { get; set; }
        public string ARM { get; set; }
        public string ARMCD { get; set; }
    }
    public class RandSubj1BllDll
    {
        public string AssgnRand2(string connectionString, string uid, string spkey, RandSubj1 randSubjInfo, string amarexDbConnStr)
        {
            var retVal = "OK";
            var chkRtnSP = "";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "BLLVpeSelUpdtRand";
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlParameter pUID = cmd.CreateParameter();
                    pUID.ParameterName = "@USERID";
                    pUID.Value = uid;
                    cmd.Parameters.Add(pUID);
                    SqlParameter pSITEID = cmd.CreateParameter();
                    pSITEID.ParameterName = "@SITEID";
                    pSITEID.Value = randSubjInfo.SITEID;
                    cmd.Parameters.Add(pSITEID);
                    SqlParameter pSUBJID = cmd.CreateParameter();
                    pSUBJID.ParameterName = "@SUBJID";
                    pSUBJID.Value = randSubjInfo.SUBJID;
                    cmd.Parameters.Add(pSUBJID);
                    SqlParameter pBRTHDTC = cmd.CreateParameter();
                    pBRTHDTC.ParameterName = "@BRTHDTC";
                    pBRTHDTC.Value = randSubjInfo.BRTHDTC;
                    cmd.Parameters.Add(pBRTHDTC);
                    SqlParameter pSEX = cmd.CreateParameter();
                    pSEX.ParameterName = "@SEX";
                    pSEX.Value = randSubjInfo.SEX;
                    cmd.Parameters.Add(pSEX);
                    SqlParameter pICDTC = cmd.CreateParameter();
                    //pICDTC.ParameterName = "@ICDTC";
                    //pICDTC.Value = randSubjInfo.ICDTCstr;
                    //cmd.Parameters.Add(pICDTC);
                    if (randSubjInfo.DateSponsorApproved == null)
                    {
                        pICDTC.ParameterName = "@ICDTC";
                        pICDTC.Value = "";
                        cmd.Parameters.Add(pICDTC);
                    }
                    else
                    {
                        pICDTC.ParameterName = "@ICDTC";
                        pICDTC.Value = randSubjInfo.ICDTCstr;
                        cmd.Parameters.Add(pICDTC);
                    }
                    SqlParameter AgeGroup = cmd.CreateParameter();
                    AgeGroup.ParameterName = "@AgeGroup";
                    AgeGroup.Value = randSubjInfo.AgeGroup;
                    cmd.Parameters.Add(AgeGroup);
                    SqlParameter ELIGRAND = cmd.CreateParameter();
                    ELIGRAND.ParameterName = "@ELIGRAND";
                    ELIGRAND.Value = randSubjInfo.ELIGRAND;
                    cmd.Parameters.Add(ELIGRAND);
                    SqlParameter StratumCode = cmd.CreateParameter();
                    StratumCode.ParameterName = "@StratumCode";
                    StratumCode.Value = randSubjInfo.StratumCode;
                    cmd.Parameters.Add(StratumCode);
                    SqlParameter pSPKEY = cmd.CreateParameter();
                    pSPKEY.ParameterName = "@SPKEY";
                    pSPKEY.Value = spkey;
                    cmd.Parameters.Add(pSPKEY);
                    SqlParameter ROW_KEY = cmd.CreateParameter();
                    ROW_KEY.ParameterName = "@ROW_KEY";
                    ROW_KEY.Value = randSubjInfo.ROW_KEY;
                    cmd.Parameters.Add(ROW_KEY);
                    SqlParameter DateSponsorApproved = cmd.CreateParameter();
                    if(randSubjInfo.DateSponsorApproved == null)
                    {
                        DateSponsorApproved.ParameterName = "@DateSponsorApproved";
                        DateSponsorApproved.Value = "";
                        cmd.Parameters.Add(DateSponsorApproved);
                    }
                    else
                    {
                        DateSponsorApproved.ParameterName = "@DateSponsorApproved";
                        DateSponsorApproved.Value = randSubjInfo.DateSponsorApproved;
                        cmd.Parameters.Add(DateSponsorApproved);
                    }
                    SqlParameter pRV = cmd.CreateParameter();
                    pRV.ParameterName = "@RV";
                    cmd.Parameters.Add(pRV);
                    cmd.Parameters["@RV"].Direction = ParameterDirection.ReturnValue;
                    int rtnSP;
                    rtnSP = cmd.ExecuteNonQuery();
                    chkRtnSP = cmd.Parameters["@RV"].Value.ToString();
                }
            }
            switch (chkRtnSP)
            {
                case "111":
                    retVal = "No Treatment Available. Please contact your CRA for questions.";
                    break;
                case "112":
                    retVal = "Kindly acknowledge the receipt of kits prior to randomization.";
                    break;
                case "223":
                    retVal = "Error Inserting Subject Information";
                    break;
                case "224":
                    retVal = "Error Updating Rand Information";
                    break;
                case "225":
                    retVal = "Error Updating IP Information";
                    break;
                case "226":
                    retVal = "Error Finding/Updating Subject and IP Information";
                    break;
                    //case "333":
                    //    retVal = "Error Dup Check";
                    //    break;
            }
            if (retVal == "OK")
            {
                var retVal2 = "";
                var othEmail = "NF";
                var retSite = "";
                GenAct genAct = new GenAct();
                int SPKEY = int.Parse(spkey);
                retVal2 = genAct.GetNotify(amarexDbConnStr, spkey, "Randomized");
                othEmail = GetProfVpe(connectionString, SPKEY, "AddRandEmails");
                retSite = genAct.GetEmailByGrp(amarexDbConnStr, spkey, randSubjInfo.SITEID, "S");
                string useremail = GetUserEmail(amarexDbConnStr, uid);
                string Study = StudyName(amarexDbConnStr, SPKEY);
                var subject = Study + " - WebView RTSM - Subject ID: " + randSubjInfo.SUBJID + " - Randomized";
                var msgBody = "";
                msgBody = "Protocol: " + Study + Environment.NewLine;
                msgBody += Environment.NewLine + "This email is to notify you that the following subject has been randomized, details below." + Environment.NewLine;
                msgBody += "Site: " + randSubjInfo.SITEID + Environment.NewLine;
                msgBody += "Subject: " + randSubjInfo.SUBJID + Environment.NewLine;
                msgBody += "Year of Birth: " + randSubjInfo.BRTHDTC + Environment.NewLine;
                msgBody += "Sex: " + randSubjInfo.SEX + Environment.NewLine;
                msgBody += "Informed consent date: " + randSubjInfo.ICDTCstr + Environment.NewLine;
                msgBody += "Age Group: " + randSubjInfo.AgeGroup + Environment.NewLine;
                var sqlState = "";
                sqlState = "SELECT * FROM [BIL_IP_RANGE] WHERE ([ROW_KEY] = " + randSubjInfo.ROW_KEY + ") ORDER BY [KitNumber]";
                DataTable dt2;
                dt2 = GetTblSql(connectionString, sqlState);
                if (dt2.Rows.Count > 0)
                {
                    msgBody += "Randomized by: " + dt2.Rows[0]["ASSIGNED_BY"].ToString() + Environment.NewLine;
                    msgBody += "Randomized date: " + (DateTime)dt2.Rows[0]["ASSIGNMENT_DATE"] + Environment.NewLine;
                    msgBody += "Visit: " + dt2.Rows[0]["VISIT"].ToString() + Environment.NewLine;
                    msgBody += "KitNumber: " + dt2.Rows[0]["KitNumber"].ToString() + Environment.NewLine;
    
                    msgBody += "ExpiryDate: " + dt2.Rows[0]["IPLblShipExpiryDate"].ToString() + Environment.NewLine;
                    msgBody += "LotNo: " + dt2.Rows[0]["IPLblShipLotNo"].ToString() + Environment.NewLine;

                }
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
                genAct.SendMail(connectionString, retVal2 + ";" + useremail, subject, msgBody);
                var arm = "";
                arm = GetARM(connectionString, randSubjInfo.SPKEY.ToString(), randSubjInfo.ROW_KEY.ToString());
                RecVisit(connectionString, randSubjInfo.SPKEY.ToString(), randSubjInfo.ROW_KEY.ToString(), randSubjInfo.SITEID, randSubjInfo.SUBJID, uid, "Visit 2", dt2.Rows[0]["KitNumber"].ToString(), arm, "Yes", "");

                string val = CheckAutoResupply(connectionString, randSubjInfo.SPKEY, randSubjInfo.SITEID);
                if ((val == "") && CheckStudyAutoRe(connectionString, randSubjInfo.SPKEY) == "Enabled")
                {
                    string result = ChkSiteInv(connectionString, randSubjInfo.SPKEY, randSubjInfo.SITEID, arm);
                    return retVal;
                }
                else
                {
                    return retVal;
                }
            }

            return retVal;

        }
        public  List<RandSubj1> GetScreenBIL(string connectionString, string spkey)
        {
            List<RandSubj1> screenList = new List<RandSubj1>();
            SqlConnection con = new SqlConnection(connectionString);
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    var sqlState = "";
                    sqlState = "SELECT * from [BIL_SUBJ] WHERE (STATUS_INFO = 'Screened' OR STATUS_INFO = 'Randomized') AND SPKEY = " + spkey + " ORDER BY SUBJID, ROW_KEY";
                    cmd.CommandText = sqlState;
                    cmd.CommandType = CommandType.Text;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            RandSubj1 getscrn = new RandSubj1();
                            getscrn.ROW_KEY = Convert.ToInt32(reader["ROW_KEY"]);
                            getscrn.SPKEY = Convert.ToInt32(reader["SPKEY"]);
                            getscrn.STATUS_INFO = reader["STATUS_INFO"].ToString();
                            getscrn.SITEID = reader["SITEID"].ToString();
                            getscrn.SUBJID = reader["SUBJID"].ToString();
                            getscrn.BRTHDTC = reader["BRTHDTC"].ToString();
                            getscrn.SEX = reader["SEX"].ToString();
                            getscrn.ICDTCstr = reader["ICDTC"].ToString();
                           // getscrn.SCRNDTCstr = reader["SCRNDTC"].ToString();
                            getscrn.DATERAND = reader["DATE_RAND"].ToString();
                            getscrn.RANDBY = reader["RANDBY"].ToString();
                            getscrn.ARMCD = reader["ARMCD"].ToString();
                            getscrn.ARM = reader["ARM"].ToString();
                            screenList.Add(getscrn);
                        }
                    }
                }
            }
            return screenList;
        }

        public static RandSubj1 GetSubjToRandBIL(string connectionString, string rowkey)
        {
            RandSubj1 getToRand = new RandSubj1();
            SqlConnection con = new SqlConnection(connectionString);
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    var sqlState = "";
                    sqlState = "SELECT * from [BIL_SUBJ] WHERE (ROW_KEY = " + rowkey + ")";
                    cmd.CommandText = sqlState;
                    cmd.CommandType = CommandType.Text;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            getToRand.SPKEY = Convert.ToInt32(reader["SPKEY"]);
                            getToRand.ROW_KEY = Convert.ToInt32(reader["ROW_KEY"]);
                            getToRand.STATUS_INFO = reader["STATUS_INFO"].ToString();
                            getToRand.SITEID = reader["SITEID"].ToString();
                            getToRand.SUBJID = reader["SUBJID"].ToString();
                            getToRand.ICDTCstr = reader["ICDTC"].ToString();
                            if (reader["STATUS_INFO"].ToString() == "Randomized")
                            {
                                getToRand.BRTHDTC = reader["BRTHDTC"].ToString();
                                getToRand.SEX = reader["SEX"].ToString();
                                getToRand.AgeGroup = reader["AgeGroup"].ToString();
                                getToRand.SCRNDTCstr = reader["SCRNDTC"].ToString();
                                getToRand.ELIGRAND = reader["ELIGRAND"].ToString();
                                getToRand.RANDBY = reader["RANDBY"].ToString();
                                getToRand.DATERAND = reader["DATE_RAND"].ToString();
                                getToRand.DateSponsorApproved = reader["DateSponsorApproved"].ToString();
                                getToRand.StratumCode = reader["StratumCode"].ToString();
                                getToRand.ARM = reader["ARM"].ToString();
                                getToRand.ARMCD = reader["ARMCD"].ToString();
                            }
                        }
                    }
                }
            }
            return getToRand;
        }
        public string ChkValEntryBIL(string connectionString, string spkey, string siteid, string subjid, string yob, string sex)
        {
            var retVal = "NV";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    int valChk = 0;
                    var sqlState = "";
                    sqlState = "SELECT COUNT(*) AS CHKDUP FROM BIL_SUBJ WHERE (SITEID = '" + siteid;
                    sqlState += "') AND ([STATUS_INFO] = 'Screened') AND (SUBJID = '" + subjid + "') AND (BRTHDTC = '" + yob + "') AND (SEX = '" + sex + "') AND (SPKEY = " + spkey + ")";
                    cmd.CommandText = sqlState;
                    cmd.CommandType = CommandType.Text;
                    valChk = (int)cmd.ExecuteScalar();
                    if (valChk == 1)
                    {
                        retVal = "OK";
                    }
                }
            }
            return retVal;
        }

        public string ChkVal(string connectionString, string spkey, string siteid, string subjid, string yob, string sex)
        {
            string sqlState = "";
            var retVal = "NV";
            string brthdtc;
            string gender;
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                
                sqlState = "SELECT BRTHDTC, SEX FROM BIL_SUBJ WHERE (SITEID = '" + siteid;
                sqlState += "') AND ([STATUS_INFO] = 'Screened') AND (SUBJID = '" + subjid + "') AND (SPKEY = " + spkey + ")";
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
                        if(sex != gender)
                        {
                            retVal = "Sex is incorrect for subject " + subjid + ".";
                            return retVal;
                        }
                        
                    }
                    else
                    {
                        // Handle the case where the SUBJID is not found in BIL_SUBJ
                       retVal = "Subject ID " + subjid + "is not valid";
                        return retVal;
                    }
                }

            }
            return retVal;
        }

        public string RecVisit(string connectionString, string spkey, string rowkey, string siteid, string subjid, string uid, string visitid, string kitno, string arm, string eligIP, string reaNo)
        {
            var rtnVal = "";
            SqlConnection con = new SqlConnection(connectionString);
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                DateTime currentDateTime = DateTime.Now;

                // Format the date as "01/Jan/2024"
                string formattedDate = currentDateTime.ToString("dd/MMM/yyyy");
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    var sqlState = "";
                    if (reaNo == "")
                    {
                        sqlState = "INSERT INTO [BIL_VISITS] ([SPKEY], [ROW_KEY], [SITEID], SUBJID, VISITID, KitNumber, ARM, ELIGIP, [ADDUSER], ADDDATE, VisitDate) VALUES (";
                        sqlState += spkey + ", " + rowkey + ", '" + siteid + "', '" + subjid + "', '" + visitid + "', '" + kitno + "', '" + arm + "', '" + eligIP + "', '" + uid + "', SYSDATETIME(), '" + formattedDate + "')";
                    }
                    else
                    {
                        sqlState = "INSERT INTO [BIL_VISITS] ([SPKEY], [ROW_KEY], [SITEID], SUBJID, VISITID, KitNumber, ARM, ELIGIP, ReaNo, [ADDUSER], ADDDATE) VALUES (";
                        sqlState += spkey + ", " + rowkey + ", '" + siteid + "', '" + subjid + "', '" + visitid + "', '" + kitno + "', '" + arm + "', '" + eligIP + "', '" + reaNo + "', '" + uid + "', SYSDATETIME())";
                    }
                    cmd.CommandText = sqlState;
                    cmd.CommandType = CommandType.Text;
                    
                    int rowsAfft = (int)cmd.ExecuteNonQuery();
                    if (rowsAfft <= 0)
                    {
                        rtnVal = "Error INSERT Rec Visit";
                    }
                }
            }
            return rtnVal;
        }

        public string GetARM(string connectionString, string spkey, string rowkey)
        {
            var rtnVal = "NF";
            SqlConnection con = new SqlConnection(connectionString);
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    var sqlState = "";
                    sqlState = "SELECT * FROM BIL_SUBJ WHERE ([ROW_KEY] = " + rowkey + ") AND (SPKey = " + spkey + ")";
                    cmd.CommandText = sqlState;
                    cmd.CommandType = CommandType.Text;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            rtnVal = reader["ARMCD"].ToString();
                        }
                    }
                }
            }
                return rtnVal;
        }

        public DataTable GetTblSql(string connectionString, string sqlState)
        {
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sqlState))
                {
                    cmd.Connection = con;
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(dt);
                    }
                }
            }
            return dt;
        }


        public string ChkSiteInv(string connectionString, int spkey, string siteid, string arm)
        {
            string rtnVal = "";
            try
            {
                string sqlState;
                sqlState = "SELECT COUNT(*) AS CHKKITS FROM BIL_IP_RANGE WHERE (SITEID = '" + siteid + "') AND (TreatmentGroup = '" + arm + "') AND (ASSIGNED IS NULL) AND (KITSTAT = 'Acceptable') AND (RECVDBY IS NOT NULL)";
                SqlConnection con = new SqlConnection(connectionString);
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
                            retSupp = GetEmailByGrp(connectionString, spkey, "(All)", "D");
                            if (retSupp == "")
                            {
                                SendEmail(connectionString,"jacobk@amarexcro.com", "Webview RTSM - Kit Shipment - Site " + siteid + " - Low Inv with Shipped status - No Supp", msgBody);
                            }
                            else
                            {
                                SendEmail(connectionString, retSupp + ";" + "jacobk@amarexcro.com", "Webview RTSM - Kit Shipment - Site " + siteid + " - Low Inv with Shipped status", msgBody);
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
                            SendEmail(connectionString,"jacobk@amarexcro.com", "Webview RTSM - Auto Re-supply - Site " + siteid + " -  Unable to find kits", "Unable to find kits for Treatment Group: " + arm);
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
                                SendEmail(connectionString,"jacobk@amarexcro.com", "Webview RTSM - Auto Re-supply - Site " + siteid + " -  Shipment process, unable to find kits", "Shipment process, unable to find kits after selection for Treatment Group: " + arm);
                            }
                            else
                            {
                                sqlState = "Auto Re-supply Shipment" + Environment.NewLine + Environment.NewLine + kitsSel;
                                rtnVal = SentKitShipEmail(connectionString, spkey, siteid, sqlState);
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


        public string SentKitShipEmail(string connectionString, int spkey, string siteid, string kits)
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
                shipperEmail = GetProfVpe(connectionString, spkey, "PMIPLblRel");
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
                    msgBody += GetShipTo(connectionString, spkey, siteid);
                    msgBody += "KitNumber in shipment: " + Environment.NewLine + kits + Environment.NewLine + Environment.NewLine;
                    if (toEmail == "")
                    {
                        toEmail = "jacobk@amarexcro.com";
                    }

                    SendEmail(connectionString, toEmail + ";" + "sidran@amarecro.com", "Webview RTSM - Request to Release - Site " + siteid, msgBody);
                }
            }
            catch (Exception exp)
            {
                throw exp;
            }
            return rtnVal;
        }

        public string GetShipTo(string connectionString, int SPKEY, string SiteID)
        {
            string retVal = "No Ship To Info Found.";
            SqlConnection con = new SqlConnection(connectionString);
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

        public string CheckAutoResupply(string connectionString, int SPKEY, string SITEID)
        {
            string rtnVal = "";
           
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

        public string CheckStudyAutoRe(string connectionString, int SPKEY)
        {
            string rtnVal = "";
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
        public string GetProfVpe(string connectionString, int SPKEY, string type)
        {
            SqlConnection con = new SqlConnection(connectionString);
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


        public string GetEmailByGrp(string connectionString, int spkey, string typesite, string grptype)
        {
            var rtnVal = "";
            using (SqlConnection conn = new SqlConnection(connectionString))
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

        public void SendEmail(string connectionString, string SendTo, string Subject, string message)
        {
            
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

        public string GetUserEmail(string connectionString, string uid)
        {

            string email = "";
            string sql = "SELECT User_Email FROM zSecurityID WHERE (UserID = '" + uid + " ')";
            SqlConnection conn = new SqlConnection(connectionString);
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

        public string StudyName(string connectionString, int SPKEY)
        {
            string rtnVal = "";
            string sql2 = "SELECT StudyName FROM STUDY_PROFILE2 WHERE  SPKEY = " + SPKEY + "";
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


    }
}
