
using AppUidAuth;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace RTSM_OLSingleArm.Controllers
{
    public class SecSSO
    {
        public string ChkIDPWSSO(string username, string password, string uriSSIS, string instanceID, string securityKey, string connectionString)
        {
            var rtnVal = "";
            GenAct chkSSO = new GenAct();
            var ssoID = "";
            ssoID = chkSSO.GetSSOuserid(connectionString, username);
            SecSS01cls chkSSOPW = new SecSS01cls();
            var encPW = "";
            encPW = chkSSOPW.EncryptRM(password, securityKey);
            rtnVal = RunAsyncUID(ssoID, encPW, uriSSIS, instanceID).GetAwaiter().GetResult();
            return rtnVal;
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
                    //rtnVal = userInfo.indicatorCode.ToString() + " - Des: " + userInfo.indicatorCodeDes.ToString();
                    rtnVal = userInfo.indicatorCode.ToString();
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
