using AppUidAuth;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace RTSM_OLSingleArm.Controllers
{
    public class CreateSesSSController : Controller
    {
        public string connectionString;
        public string setUidSess;
        readonly IConfiguration _configuration;
        public CreateSesSSController(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public IActionResult Index(string SPKey, string userToken)
        {
            HttpContext.Session.SetString("userToken", userToken);
            HttpContext.Session.SetString("sesSPKey", SPKey);
            HttpContext.Session.SetString("CreateSesSS", "CreateSesSS");
            connectionString = _configuration.GetConnectionString("AmarexDbConnStr");
            var connectionStringVPE = _configuration.GetConnectionString("VpeRandDbConnStr");
            var connectionDB = _configuration.GetConnectionString("AmarexDb");
            HttpContext.Session.SetString("sesAmarexDb", connectionDB);
            HttpContext.Session.SetString("sesAmarexDbConnStr", connectionString);
            HttpContext.Session.SetString("sesVpeRandDbConnStr", connectionStringVPE);
            var uriSSIS = "";
            uriSSIS = _configuration["appSettings:uriSSIS"];
            HttpContext.Session.SetString("sesuriSSIS", uriSSIS);
            var instanceID = "";
            instanceID = _configuration["appSettings:instanceID"];
            HttpContext.Session.SetString("sesinstanceID", instanceID);
            var securityKey = "";
            securityKey = _configuration["appSettings:SecurityKey"];
            HttpContext.Session.SetString("sesSecurityKey", securityKey);
            RunAsync(userToken, uriSSIS, instanceID, securityKey, HttpContext.Session.GetString("sesAmarexDbConnStr"), HttpContext.Session).GetAwaiter().GetResult();
            if (HttpContext.Session.Get("suserid") != null)
            {
                if (HttpContext.Session.GetString("suserid") == "Signin Error")
                {
                    return RedirectToAction("Privacy", "Home");
                }
            }
            //return View();
            AppUidClass2 auc2 = new AppUidClass2(HttpContext.Session);
            var rtnVal = "";
            var rtnValSite = "";
            //Method to created role Sessions
            rtnVal = auc2.GetUIDRole2(connectionString, HttpContext.Session.GetString("suserid"));
            //Method to created site Session
            rtnValSite = auc2.GetUIDCtr(connectionString, SPKey, HttpContext.Session.GetString("suserid"));
            if (rtnValSite == "OK")
            {
                //HttpContext.Session.SetString("suserid", HttpContext.Session.GetString("suserid"));
                HttpContext.Session.SetString("sesSPKey", SPKey);
            }
            return RedirectToAction("Index", "Home");
        }

        static async Task RunAsync(string userToken, string uriSSIS, string instanceID, string securityKey, string connectionString, ISession session)
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(uriSSIS);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("instanceID", instanceID);
                client.DefaultRequestHeaders.Add("userToken", userToken);
                var responseTask = client.GetAsync("user");
                var result = responseTask.Result;
                MySessionWrapper setses = new MySessionWrapper(session);
                SecProc secinfo = new SecProc();
                if (result.IsSuccessStatusCode)
                {
                    var userInfo = await JsonSerializer.DeserializeAsync<GetUserInfo>(await result.Content.ReadAsStreamAsync());
                    var chkuser = userInfo.indicatorCodeDes.ToString();
                    if (userInfo.indicatorCode.ToString() == "7201")
                    {
                        string rtnVal = "";
                        rtnVal = DecryptRM(userInfo.user.ToString(), securityKey);
                        setses.SetUid(rtnVal);
                        rtnVal = secinfo.InsSignInLog(connectionString, rtnVal, "Sucessful Login", "SSIS");
                    }
                    else
                    {
                        setses.SetUid("Signin Error");
                        secinfo.InsSignInLog(connectionString, userToken, userInfo.indicatorCodeDes.ToString(), "SSIS");
                    }
                }
                else
                {
                    setses.SetUid("Signin Error");
                    secinfo.InsSignInLog(connectionString, userToken, result.StatusCode.ToString(), "SSIS");
                }
            }
        }
        public class GetUserInfo
        {
            public string indicatorCode { get; set; }
            public string indicatorCodeDes { get; set; }
            public string user { get; set; }
        }
        public static string DecryptRM(string toDecrypt, string key)
        {
            byte[] keyArray = System.Text.UTF8Encoding.UTF8.GetBytes(key);
            byte[] toEncryptArray = Convert.FromBase64String(toDecrypt);

            RijndaelManaged rDel = new RijndaelManaged();
            rDel.Key = keyArray;
            rDel.Mode = CipherMode.ECB;
            rDel.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = rDel.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);

            return UTF8Encoding.UTF8.GetString(resultArray);
        }

        public class MySessionWrapper
        {
            ISession session;
            public MySessionWrapper(ISession session)
            {
                this.session = session;
            }
            public void SetUid(string uid)
            {
                session?.SetString("suserid", uid);
            }
        }
    }
}
