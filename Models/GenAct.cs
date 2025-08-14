using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class xGenAct
    {
        public string vpeRandDbConnStr;
        public string amarexDbConnStr;

        public string GetStudyName(string connectionString, string spkey)
        {
            var studyName = "Error - Name Not Found";
            //amarexDbConnStr = _configuration.GetConnectionString("AmarexDbConnStr");
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "GenSelSPNameByKey";
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlParameter param = cmd.CreateParameter();
                    param.ParameterName = "@pkey";
                    param.Value = spkey;
                    cmd.Parameters.Add(param);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            studyName = reader["StudyName"].ToString();
                        }
                    }
                }
            }
            return studyName;
        }

        public string GetSSOuserid(string connectionString, string uid)
        {
            var ssouid = "Error - ID Not Found";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SecSelUserID";
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlParameter param = cmd.CreateParameter();
                    param.ParameterName = "@iuserid";
                    param.Value = uid;
                    cmd.Parameters.Add(param);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            ssouid = reader["SSO_user_id"].ToString();
                        }
                    }
                }
            }
            return ssouid;
        }
        public string GetNotify(string connectionString, string spkey, string notifyFor)
        {
            var rtnVal = "";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    var sqlState = "";
                    sqlState = "SELECT zSecurityID.User_Email FROM zSecurityID INNER JOIN NOTIFY_BY_STUDY ";
                    sqlState += "ON zSecurityID.UserID = NOTIFY_BY_STUDY.USERID WHERE (NOTIFY_BY_STUDY.SPKEY = " + spkey;
                    sqlState += ") AND (NOTIFY_BY_STUDY.NOTIFY_FOR = '" + notifyFor + "')";
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
        public string GetEmailByGrp(string connectionString, string spkey, string typesite, string grptype)
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
        public string GetProfVpe(string connectionString, string spkey, string profType)
        {
            var rtnVal = "NF";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "GenSelPIBySPType";
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlParameter pSPKEY = cmd.CreateParameter();
                    pSPKEY.ParameterName = "@SPKEY";
                    pSPKEY.Value = spkey;
                    cmd.Parameters.Add(pSPKEY);
                    SqlParameter pType = cmd.CreateParameter();
                    pType.ParameterName = "@PIType";
                    pType.Value = profType;
                    cmd.Parameters.Add(pType);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            rtnVal = reader["PIDet"].ToString();
                        }
                    }
                }
            }
            return rtnVal;
        }
        public void SendMail(string connectionString, string SENDMAILTO, string SUBJ, string MSGBODY)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "NotifySend2";
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlParameter pSENDMAILTO = cmd.CreateParameter();
                    pSENDMAILTO.ParameterName = "@SENDMAILTO";
                    pSENDMAILTO.Value = SENDMAILTO;
                    cmd.Parameters.Add(pSENDMAILTO);
                    SqlParameter pSUBJ = cmd.CreateParameter();
                    pSUBJ.ParameterName = "@SUBJ";
                    pSUBJ.Value = SUBJ;
                    cmd.Parameters.Add(pSUBJ);
                    SqlParameter pMSGBODY = cmd.CreateParameter();
                    pMSGBODY.ParameterName = "@MSGBODY";
                    pMSGBODY.Value = MSGBODY;
                    cmd.Parameters.Add(pMSGBODY);
                    cmd.ExecuteNonQuery();
                }
            }
        }
        public void InsSignInLog(string connectionString, string sUserID, string sMsg, string sIP_ADDR)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SecInsSignInLog3";
                    cmd.CommandType = CommandType.StoredProcedure;
                    SqlParameter pUID = cmd.CreateParameter();
                    pUID.ParameterName = "@UserID";
                    pUID.Value = sUserID;
                    cmd.Parameters.Add(pUID);
                    SqlParameter pMsg = cmd.CreateParameter();
                    pMsg.ParameterName = "@DescMSG";
                    pMsg.Value = sMsg;
                    cmd.Parameters.Add(pMsg);
                    SqlParameter pIP_ADDR = cmd.CreateParameter();
                    pIP_ADDR.ParameterName = "@IP_ADDR";
                    pIP_ADDR.Value = sIP_ADDR;
                    cmd.Parameters.Add(pIP_ADDR);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public async Task<string> RunAsyncUID(string uid, string pw, string uriSSIS, string instanceID)
        {
            VerPW model = new VerPW();
            string rtnVal = "Error";
            using (var client = new HttpClient())
            {
                //string uriSSIS = System.Configuration.ConfigurationManager.AppSettings["uriSSIS"];
                client.BaseAddress = new Uri(uriSSIS);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                var bodyVerPW = new VerPWBodyUID();
                //string instanceID = System.Configuration.ConfigurationManager.AppSettings["InstanceID"];
                bodyVerPW.instanceID = instanceID;
                bodyVerPW.userID = uid;
                bodyVerPW.password = pw;
                var responseTask = client.PostAsJsonAsync("verifyPasswordByID", bodyVerPW);
                var result = responseTask.Result;
                if (result.IsSuccessStatusCode)
                {
                    var userInfo = await result.Content.ReadAsAsync<VerPW>();
                    var chkuser = userInfo.indicatorCodeDes.ToString();
                    //rtnVal = userInfo.indicatorCodeDes.ToString();
                    //rtnVal = userInfo.indicatorCode.ToString();
                    rtnVal = userInfo.indicatorCode.ToString() + " - Des: " + userInfo.indicatorCodeDes.ToString();
                }
                else
                {
                    rtnVal = result.StatusCode.ToString();
                }
            }
            return rtnVal;
        }
        public class VerPWBody
        {
            public string instanceID { get; set; }
            public string username { get; set; }
            public string password { get; set; }
        }
        public class VerPWBodyUID
        {
            public string instanceID { get; set; }
            public string userID { get; set; }
            public string password { get; set; }
        }
        public class VerPW
        {
            public string indicatorCode { get; set; }
            public string indicatorCodeDes { get; set; }
        }




    }
}
