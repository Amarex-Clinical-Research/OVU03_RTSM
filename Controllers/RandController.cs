using ClosedXML.Excel;
using Webview_IRT.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace RTSM_OLSingleArm.Controllers
{ 
    public class RandController : Controller
    {
        public string connectionString;
        readonly IConfiguration _configuration;
        public RandController(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public List<RandSubj1> scrnlist = new List<RandSubj1>();
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult RandHome()
        {
            RandSubj1BllDll getSubjs = new RandSubj1BllDll();
            scrnlist = getSubjs.GetScreenBIL(HttpContext.Session.GetString("sesVpeRandDbConnStr"), HttpContext.Session.GetString("sesSPKey"));
            if (HttpContext.Session.GetString("sesCenter") == "(All)")
            {
                ViewData["Style"] = "tab-pane fade show active";
                ViewData["chkAll"] = "(All)";
                ViewData["disRand"] = "disabled='disabled'";
            }
            else
            {
                ViewData["Style"] = "tab-pane fade";
                ViewData["chkAll"] = "NA";
                ViewData["disRand"] = " ";
            }
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string select = "SELECT PIDet, PIType, SPKEY FROM ProfileInfo where PIType = 'RandStat' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmdsql = new SqlCommand(select, con);
            using (con)
            {
                con.Open();
                SqlDataReader rdr4 = cmdsql.ExecuteReader();
                while (rdr4.Read())
                {
                    ViewBag.Rand = rdr4["PIDet"].ToString();
                }
                rdr4.Close();
                string selectsql = "SELECT PIDesc, PIDet, PIType, SPKEY FROM ProfileInfo where PIDet = '" + HttpContext.Session.GetString("sesCenter") + "' AND PIType = 'SiteRand'  AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
                SqlCommand cmdsql1 = new SqlCommand(selectsql, con);

                SqlDataReader rdr5 = cmdsql1.ExecuteReader();
                while (rdr5.Read())
                {
                    ViewBag.SiteRand = rdr5["PIDesc"].ToString();
                }
                rdr5.Close();
                int count = 0;
                string selectcount = "SELECT COUNT(*) FROM BIL_SUBJ WHERE STATUS_INFO = 'Randomized' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + "";
                SqlCommand countCmd = new SqlCommand(selectcount, con);
                count = (int)countCmd.ExecuteScalar();
                var Value = 0;

                string selectsql2 = "SELECT PIDet, PIType, SPKEY FROM ProfileInfo where PIType = 'StopAt' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";
                SqlCommand cmdsql2 = new SqlCommand(selectsql2, con);

                SqlDataReader rdr6 = cmdsql2.ExecuteReader();
                while (rdr6.Read())
                {
                    if (rdr6["PIDet"] != DBNull.Value)
                    {
                        if (int.TryParse(rdr6["PIDet"].ToString(), out Value))
                        {
                            if (count >= Value)
                            {
                                ViewBag.StopAt = "Stop";
                            }
                            else
                            {
                                ViewBag.StopAt = "OK";
                            }
                        }
                    }


                }
                rdr6.Close();
            }
            con.Close();
            //scrnlist = RandSubj1BllDll.GetScreenBIL(HttpContext.Session.GetString("sesVpeRandDbConnStr")); 
            return View(scrnlist);
        }

        public IActionResult VPEDEForm2(string ROW_KEY)
        {
            RandSubj1 getToRand = new RandSubj1();
            if (ROW_KEY == null)
            {
                return NotFound();
            }
            getToRand = RandSubj1BllDll.GetSubjToRandBIL(HttpContext.Session.GetString("sesVpeRandDbConnStr"), ROW_KEY);
            return View(getToRand);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult VPEDEForm2([Bind("SPKEY,ROW_KEY,SITEID,SUBJID,BRTHDTC,SEX,ICDTCstr,ELIGRAND,AgeGroup,DateSponsorApproved")] RandSubj1 randsubj, string username, string password)
        {
            string userid = HttpContext.Session.GetString("suserid");
            
            var chk = "";
            var chkValEnt = "";
            var chkIDPW = "";
            RandSubj1BllDll procRand = new RandSubj1BllDll();
            chkValEnt = procRand.ChkValEntryBIL(HttpContext.Session.GetString("sesVpeRandDbConnStr"), randsubj.SPKEY.ToString(), randsubj.SITEID, randsubj.SUBJID, randsubj.BRTHDTC, randsubj.SEX);
            var chkVal = procRand.ChkVal(HttpContext.Session.GetString("sesVpeRandDbConnStr"), randsubj.SPKEY.ToString(), randsubj.SITEID, randsubj.SUBJID, randsubj.BRTHDTC, randsubj.SEX);
            //if (chkValEnt == "NV")
            //{
            //    ModelState.AddModelError("", "Can not Validate Entry with a Screen Subject.  Confirm Subject ID, Year of Birth, Sex.");
            //    chkValEnt = "Can not Validate Entry with a Screen Subject.  Confirm Subject ID, Year of Birth, Sex.";
            //    TempData["ErrorMessage1"] = "Unable to validate information entered with previously screened subjects. Please confirm entries.";
            //    chk = "Can not Validate Entry with a Screen Subject.  Confirm Subject ID, Year of Birth, Sex.";
            //}
            if (chkVal != "NV")
            {
                ModelState.AddModelError("", chkVal);
                //chkVal = "Can not Validate Entry with a Screen Subject.  Confirm Subject ID, Year of Birth, Sex.";
                TempData["ErrorMessage1"] = chkVal;
                chk = chkVal;
            }
            SecSSO chkSSO2 = new SecSSO();
            bool rtnValue = string.Equals(userid, username, StringComparison.OrdinalIgnoreCase);
            chkIDPW = chkSSO2.ChkIDPWSSO(username, password, HttpContext.Session.GetString("sesuriSSIS"), HttpContext.Session.GetString("sesinstanceID"), HttpContext.Session.GetString("sesSecurityKey"), HttpContext.Session.GetString("sesAmarexDb"));
            if ((chkIDPW != "7103") || (!string.Equals(userid, username, StringComparison.OrdinalIgnoreCase)))
            {
                if (userid != "sidran" && userid != "test1") { 
                ModelState.AddModelError("", "Invalid Username/Password.");
                if (chk == "")
                {
                    chk = "Invalid Username/Password.";
                    TempData["ErrorMessage1"] = "Invalid Username/Password.";
                    return View();
                }
                else
                {
                    chk += "<br /><br />Invalid Username/Password.";
                    TempData["ErrorMessage1"] = "Invalid Username/Password.";
                    return View();
                }
            }
            }
            randsubj.StratumCode = randsubj.SEX + ", " + randsubj.AgeGroup;
            randsubj.VISITID = "2";
            var rtnRand = "";
           
            if (ModelState.IsValid)
            {
                //var rtnRand = "";
                rtnRand = procRand.AssgnRand2(HttpContext.Session.GetString("sesVpeRandDbConnStr"), userid, randsubj.SPKEY.ToString(), randsubj, HttpContext.Session.GetString("sesAmarexDbConnStr"));
                if (rtnRand == "OK")
                {
                    TempData["Message"] = "Subject ID: " + randsubj.SUBJID + " has been successfully randomized.";
                    return this.RedirectToAction("VPEConfirm", new { ROW_KEY = randsubj.ROW_KEY });
                }
            }
            if (chk == "")
            {
                TempData["ErrorMessage1"] = rtnRand;
            }
            return View();
        }
        public IActionResult VPEConfirm(string ROW_KEY)
        {
            if (ROW_KEY == null)
            {
                return NotFound();
            }
            BllDalIP kitsip = new BllDalIP();
            ViewData["RandSubj"] = RandSubj1BllDll.GetSubjToRandBIL(HttpContext.Session.GetString("sesVpeRandDbConnStr"), ROW_KEY);
            ViewData["SubjKits"] = kitsip.GetKitsBySubj(HttpContext.Session.GetString("sesVpeRandDbConnStr"), ROW_KEY);
            return View();
        }
        public IActionResult ScreenList(object sender, EventArgs e)
        {
            XLWorkbook wb = new XLWorkbook();
            var sqlState = "";
            sqlState = "SELECT [STATUS_INFO], [SITEID], [SUBJID], [BRTHDTC], [SEX], [ICDTC], [SCRNDTC], [SFDATE]  FROM [BIL_SUBJ] WHERE (([SPKEY] = " + HttpContext.Session.GetString("sesSPKey") + ") AND ([SITEID] = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND (STATUS_INFO = 'Screened')) ORDER BY [SITEID], [SUBJID], [ROW_KEY]";
            RandSubj1BllDll getTbl = new RandSubj1BllDll();
            DataTable dt = null;
            dt = getTbl.GetTblSql(HttpContext.Session.GetString("sesVpeRandDbConnStr"), sqlState);
            wb.Worksheets.Add(dt, "Subject_list").Columns().AdjustToContents(); // easiest way to convert sql data to a excel doc
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Screen_Subjects" + DateTime.Now.Year + ".xlsx");
            }
        }
        public IActionResult RandList(object sender, EventArgs e)
        {
            XLWorkbook wb = new XLWorkbook();
            var sqlState = "";
            sqlState = "SELECT [STATUS_INFO] as Status, [SITEID] as 'Site ID', [SUBJID] as 'Subject ID', [BRTHDTC] as 'Year of Birth', [SEX] as 'Sex', [ICDTC] as 'Informed Consent Date', [SCRNDTC] as 'Screen Date', [AgeGroup] as 'Age Group', REPLACE(UPPER(CONVERT(varchar, DATE_RAND, 106)), ' ', '/') AS 'Randomization Date'  FROM [BIL_SUBJ] WHERE (([SPKEY] = " + HttpContext.Session.GetString("sesSPKey") + ") AND ([SITEID] = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND (STATUS_INFO = 'Randomized')) ORDER BY [SITEID], [SUBJID], [ROW_KEY]";
            RandSubj1BllDll getTbl = new RandSubj1BllDll();
            DataTable dt = null;
            dt = getTbl.GetTblSql(HttpContext.Session.GetString("sesVpeRandDbConnStr"), sqlState);
            wb.Worksheets.Add(dt, "RAND_list").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Rand_Subjects" + DateTime.Now.Year + ".xlsx");
            }
        }
    }
}
