using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using AppUidAuth;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Http;

namespace RTSM_OLSingleArm.Controllers
{
    public class CreateSesModel : Controller
    {
        //public IActionResult Index()
        //{
        //    return View();
        //}

        public string connectionString;
        readonly IConfiguration _configuration;
        public CreateSesModel(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public IActionResult Index(string SPKey, string UID)
        {
            //Get connection info from appSettings.json
            var connectionStringVPE = _configuration.GetConnectionString("VpeRandDbConnStr");
            var connectionDB = _configuration.GetConnectionString("AmarexDb");
            connectionString = _configuration.GetConnectionString("AmarexDbConnStr");
            HttpContext.Session.SetString("sesAmarexDbConnStr", connectionString);
            HttpContext.Session.SetString("sesVpeRandDbConnStr", connectionStringVPE);
            HttpContext.Session.SetString("sesAmarexDb", connectionDB);
            var uriSSIS = "";
            uriSSIS = _configuration["appSettings:uriSSIS"];
            HttpContext.Session.SetString("sesuriSSIS", uriSSIS);
            var instanceID = "";
            instanceID = _configuration["appSettings:instanceID"];
            HttpContext.Session.SetString("sesinstanceID", instanceID);
            var securityKey = "";
            securityKey = _configuration["appSettings:SecurityKey"];
            HttpContext.Session.SetString("sesSecurityKey", securityKey);
            //Create the class, send Session object
            AppUidClass2 auc2 = new AppUidClass2(HttpContext.Session);
            var rtnVal = "";
            var rtnValSite = "";
            //Method to created role Sessions
            rtnVal = auc2.GetUIDRole2(connectionString, UID);
            //Method to created site Session
            rtnValSite = auc2.GetUIDCtr(connectionString, SPKey, UID);
            if (rtnValSite == "OK")
            {
                HttpContext.Session.SetString("suserid", UID);
                HttpContext.Session.SetString("sesSPKey", SPKey);
                return RedirectToAction("Index", "Home");
            }
            return RedirectToPage("/ErrorPage");
        }

       


    }
}
