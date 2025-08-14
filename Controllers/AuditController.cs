using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Authorization;
using System.Data;
using Webview_IRT.Models;


namespace RTSM_OLSingleArm.Controllers
{
    public class AuditController : Controller
    {
        public string connectionString;
        readonly IConfiguration _configuration;
        public AuditController(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public IActionResult AuditHome()
        {
            return View();
        }
        public IActionResult AuditLog(string start, string end, string Module)
        {
            TempData["start"] = start;
            TempData["end"] = end;
            
            List<Audit> auditLog = new List<Audit>();
            string sql = "";
            if (Module == "BIL_SUBJ" || Module == "BIL_UNBLD")
                sql = "select * From TRAUD_UNBlinded where AUDTBL = '" + Module + "' AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM') AND (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE))) order by AUDDTC desc";
            else if (Module == "BIL_RAND")
                sql = "select * From TRAUD_UNBlinded where (AUDTBL = '" + Module + "' OR AUDTBL = 'BIL_SUBJ' OR AUDTBL = 'BIL_IP_RANGE' OR AUDTBL = 'BIL_VISITS') AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM')  AND (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  order by AUDDTC desc";
            else if (Module == "BIL_VISITS")
                sql = "select * From TRAUD_UNBlinded where (AUDTBL = 'BIL_IP_RANGE' OR AUDTBL = 'BIL_VISITS') AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM')  AND (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  order by AUDDTC desc";
            else if (Module == "BIL_IP_RANGE")
                sql = "select * From TRAUD_UNBlinded where AUDTBL = '" + Module + "' AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM') AND (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  order by AUDDTC desc";
            else if (Module == "IP Labeling")
            {
                sql = "SELECT * FROM (SELECT * FROM TRAUD_UNBlinded WHERE AUDTBL = 'BIL_IP_RANGE' AND (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM'))AS tbl1";
                sql += " UNION ALL ";
                sql += "SELECT * FROM TRAUD_Blinded WHERE AUDTBL IN ('IP_LABEL_SHIP', 'IP_LABEL_SHIP_REQ') AND (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM') ORDER BY AUDDTC DESC";
            }
            else if (Module == "IP Depot")
            {
                sql = "SELECT * FROM (SELECT * FROM TRAUD_UNBlinded WHERE AUDTBL = 'BIL_IP_RANGE' AND (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM')) AS tbl1";
                sql += " UNION ALL ";
                sql += "SELECT * FROM TRAUD_Blinded WHERE AUDTBL IN ('IP_LABEL_DEPOT_REL') AND (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM') ORDER BY AUDDTC DESC";
            }
            else if (Module == "Project Management")
            {
                sql = "SELECT * FROM (SELECT * FROM TRAUD_UNBlinded WHERE (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM')) AS tbl1";
                sql += " UNION ALL ";
                sql += "SELECT * FROM TRAUD_Blinded WHERE (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM') ORDER BY AUDDTC DESC";
            }
            else if (Module == "IP_BIO_UPLOADS" || Module == "BIL_PRBLM")
                sql = "select * From TRAUD_Blinded where AUDTBL = '" + Module + "' AND (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  order by AUDDTC desc";
            else if(Module == "Receive IP")
            {
                sql = "SELECT * FROM (SELECT * FROM TRAUD_UNBlinded WHERE AUDTBL = 'BIL_IP_RANGE' AND (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM')) AS tbl1";
                sql += " UNION ALL ";
                sql += "SELECT * FROM TRAUD_Blinded WHERE AUDTBL IN ('IPUploads') AND (AUDDTC between CONVERT(datetime, '" + start + "') and DATEADD(day, 1, CAST('" + end + "' AS DATE)))  AND (AUDFLD != 'TreatmentGroup' AND AUDFLD != 'TreatmentName' AND AUDFLD != 'ARMCD' AND AUDFLD != 'ARM') ORDER BY AUDDTC DESC";
            }
            else
                return RedirectToAction("AuditHome");
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    var temp = new Audit();
                    temp.AUDTYPE = rdr["AUDTYPE"].ToString();
                    temp.AUDFLD = rdr["AUDFLD"].ToString();
                    temp.AUDOLDV = rdr["AUDOLDV"].ToString();
                    temp.AUDNEWV = rdr["AUDNEWV"].ToString();
                    temp.AUDDTC = (DateTime)(rdr["AUDDTC"]); 
                    temp.AUDUSER = rdr["AUDUSER"].ToString();
                    temp.AUDTBL = rdr["AUDTBL"].ToString();
                    if (temp.AUDTBL == null || temp.AUDTBL == "")
                    {
                        temp.AUDTBL = rdr["AUDTBL"].ToString();
                    }
                    temp.AUDPKEY = rdr["AUDPKEY"].ToString();
                    auditLog.Add(temp);
                }
            }
            return View(auditLog); 
        }
    }
}
