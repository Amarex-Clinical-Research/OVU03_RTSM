using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Threading.Tasks;
using Webview_IRT.Models;

namespace RTSM_OLSingleArm.Controllers
{
    public class ReportsController : Controller
    {
        public string connectionString;
        readonly IConfiguration _configuration;
        public ReportsController(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public IActionResult ReportsHome()
        {
            ViewBag.Back = "ReportHome";
            return View();
        }

        public IActionResult SubjectReport()
        { //since the study table does not currently contain all of the codes to allow for joining I full joined all the log tables on project codes and will just take the relavent info from each
          // string sql = "select * from PROJECT_TEAM_LOG a full join ClinProjectLogs b on a.PROJECTCODE = b.ProjectCode full join DMProjectLogs c on a.PROJECTCODE = c.AmaProjCode full join ITProjectLogs d on a.PROJECTCODE = d.ProjectCode full join MWProjectLogs e on a.PROJECTCODE = e.ProjectCode full join RAProjectLogs f on a.PROJECTCODE = f.ProjectCode full join SafetyProjectLogs g on a.PROJECTCODE = g.StudyCode";
            string sql = "SELECT ARM, ARMCD, SITEID, SUBJID, BRTHDTC, SEX, AgeGroup, ICDTC,  STATUS_INFO, PMCOMDATE, ADDDATE, ADDUSER, DATE_RAND, RANDBY, SFDATE FROM BIL_SUBJ WHERE (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID   ";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Subject_Status" + DateTime.Now.ToString("MM-dd-yyyy"));


            worksheet.Cell(1, 1).Value = "Site ID";
            worksheet.Cell(1, 2).Value = "Subject ID";
            worksheet.Cell(1, 3).Value = "Year of Birth";
            worksheet.Cell(1, 4).Value = "Sex";
            worksheet.Cell(1, 5).Value = "Age Group";
            worksheet.Cell(1, 6).Value = "Informed Consent Date";
            worksheet.Cell(1, 7).Value = "Status";
            //worksheet.Cell(1, 8).Value = "Screened Date";
            worksheet.Cell(1, 8).Value = "Screened By";
            worksheet.Cell(1, 9).Value = "Randomization Date";
            worksheet.Cell(1, 10).Value = "Randomized By";
            worksheet.Cell(1, 11).Value = "Treatment";
            worksheet.Cell(1, 12).Value = "Treatment Group";
            worksheet.Cell(1, 13).Value = "Screen Failed Date";
           


            int i = 2;

            while (rdr.Read())
            {
                worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["SITEID"]) ? rdr["SITEID"].ToString() : null;
                worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["SUBJID"]) ? rdr["SUBJID"].ToString() : null;
                worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["BRTHDTC"]) ? rdr["BRTHDTC"].ToString() : null;
                worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["SEX"]) ? rdr["SEX"].ToString() : null;
                worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["AgeGroup"]) ? rdr["AgeGroup"].ToString() : null;
                worksheet.Cell(i, 6).Value = !Convert.IsDBNull(rdr["ICDTC"]) ? rdr["ICDTC"].ToString() : null;
                worksheet.Cell(i, 7).Value = !Convert.IsDBNull(rdr["STATUS_INFO"]) ? rdr["STATUS_INFO"].ToString() : null;
                //worksheet.Cell(i, 8).Value = !Convert.IsDBNull(rdr["SCRNDTC"]) ? rdr["SCRNDTC"].ToString() : null;
                worksheet.Cell(i, 8).Value = !Convert.IsDBNull(rdr["ADDUSER"]) ? rdr["ADDUSER"].ToString() : null;
                worksheet.Cell(i, 9).Value = !Convert.IsDBNull(rdr["DATE_RAND"]) ? rdr["DATE_RAND"].ToString() : null;
                worksheet.Cell(i, 10).Value = !Convert.IsDBNull(rdr["RANDBY"]) ? rdr["RANDBY"].ToString() : null;
                worksheet.Cell(i, 11).Value = !Convert.IsDBNull(rdr["ARM"]) ? rdr["ARM"].ToString() : null;
                worksheet.Cell(i, 12).Value = !Convert.IsDBNull(rdr["ARMCD"]) ? rdr["ARMCD"].ToString() : null;
                worksheet.Cell(i, 13).Value = !Convert.IsDBNull(rdr["SFDATE"]) ? rdr["SFDATE"].ToString() : null;
                i++;

            }
            conn.Close();
            worksheet.Columns().AdjustToContents();
            return View(ImportExceltoDatatable(workbook));

        }

        public static DataTable ImportExceltoDatatable(XLWorkbook file)
        {
            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (file)
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = file.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;

                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }

                return dt;
            }



        }



        public static DataTable ImportExceltoDatatable2(XLWorkbook file)
        {
            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (file)
            {
                // Read the first Sheet from Excel file.
                IXLWorksheet workSheet = file.Worksheet(1);

                // Create a new DataTable.
                DataTable dt = new DataTable();

                // Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    // Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        // Add rows to DataTable.
                        dt.Rows.Add();

                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            // Ensure the cell is not null before reading its value
                            if (row.Cell(i + 1) != null)
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = row.Cell(i + 1).Value.ToString();
                            }
                            else
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = null;
                            }
                        }
                    }
                }

                return dt;
            }
        }



        public IActionResult SubjectReportExcel()
        {

            XLWorkbook wb = new XLWorkbook();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DataTable dt = CreateDataTable("SELECT SITEID AS 'Site ID', SUBJID AS 'Subject ID', BRTHDTC AS 'Year of Birth', SEX AS 'Sex', AgeGroup as 'Age Group', ICDTC AS 'Informed Consent Date', STATUS_INFO AS 'Status', ADDUSER AS 'Screened By', DATE_RAND AS 'Randomization Date', RANDBY AS 'Randomized By', ARM AS 'Treatment', ARMCD AS 'Treatment Group', SFDATE As 'Screen Failed Date ' FROM BIL_SUBJ WHERE (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND SPKEY  = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID   ", connectionString);
            wb.Worksheets.Add(dt, "Subject Status Report").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Subject Status Report" + DateTime.Now.Year + ".xlsx");
            }

        }


        //public IActionResult SubjectVisitReport()
        //{
        //    string sql = "SELECT [VISIT], KitNumber, REPLACE(UPPER(CONVERT(varchar, [ASSIGNMENT_DATE], 106)), ' ', '/') AS 'Dispense', ASSIGNED_BY, ExpiryDate, LotNo, KitRepled, SITEID, SUBJID, notdone, ReaYes FROM BIL_IP_RANGE WHERE SITEID is not NULL AND SUBJID is not null ORDER BY SITEID, SUBJID";
        //    connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
        //    SqlConnection conn = new SqlConnection(connectionString);
        //    SqlCommand cmd = new SqlCommand(sql, conn);
        //    conn.Open();
        //    SqlDataReader rdr = cmd.ExecuteReader();

        //    var workbook = new XLWorkbook();
        //    var worksheet = workbook.Worksheets.Add("Subject_Visit" + DateTime.Now.ToString("MM-dd-yyyy"));


        //    worksheet.Cell(1, 1).Value = "Site ID";
        //    worksheet.Cell(1, 2).Value = "Subject ID";
        //    worksheet.Cell(1, 3).Value = "Visit";
        //    worksheet.Cell(1, 4).Value = "Kit Number";
        //    worksheet.Cell(1, 5).Value = "Dispensed Date";
        //    worksheet.Cell(1, 6).Value = "Dispensed By";
        //    worksheet.Cell(1, 7).Value = "Expiry Date";
        //    worksheet.Cell(1, 8).Value = "Lot Number";
        //    worksheet.Cell(1, 9).Value = "Replacement For";

        //    worksheet.Cell(1, 10).Value = "Visit Not done";
        //    worksheet.Cell(1, 11).Value = "Reason(IP dispensed, visit not done)";
        //    //worksheet.Cell(1, 15).Value = "PV Monitor";
        //    //worksheet.Cell(1, 16).Value = "Authorized Representative";


        //    int i = 2;

        //    while (rdr.Read())
        //    {
        //        worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["SITEID"]) ? rdr["SITEID"].ToString() : null;
        //        worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["SUBJID"]) ? rdr["SUBJID"].ToString() : null;
        //        worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["VISIT"]) ? rdr["VISIT"].ToString() : null;
        //        worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["KitNumber"]) ? rdr["KitNumber"].ToString() : null;
        //        worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["Dispense"]) ? rdr["Dispense"].ToString() : null;
        //        worksheet.Cell(i, 6).Value = !Convert.IsDBNull(rdr["ASSIGNED_BY"]) ? rdr["ASSIGNED_BY"].ToString() : null;
        //        worksheet.Cell(i, 7).Value = !Convert.IsDBNull(rdr["IPLblShipExpiryDate"]) ? rdr["IPLblShipExpiryDate"].ToString() : null;
        //        worksheet.Cell(i, 8).Value = !Convert.IsDBNull(rdr["IPLblShipLotNo"]) ? rdr["IPLblShipLotNo"].ToString() : null;
        //        worksheet.Cell(i, 9).Value = !Convert.IsDBNull(rdr["KitRepled"]) ? rdr["KitRepled"].ToString() : null;
        //        worksheet.Cell(i, 10).Value = !Convert.IsDBNull(rdr["notdone"]) ? rdr["notdone"].ToString() : null;
        //        worksheet.Cell(i, 11).Value = !Convert.IsDBNull(rdr["ReaYes"]) ? rdr["ReaYes"].ToString() : null;
        //        //worksheet.Cell(i, 12).Value = !Convert.IsDBNull(rdr["SFDATE"]) ? rdr["SFDATE"].ToString() : null;
        //        i++;

        //    }
        //    conn.Close();
        //    worksheet.Columns().AdjustToContents();
        //    return View(ImportExceltoDatatable(workbook));

        //}

        public IActionResult SubjectVisitReport()
        {
            string sql = "SELECT ARM, ROW_KEY, SITEID, SUBJID, BRTHDTC, SEX, AgeGroup, ICDTC, STATUS_INFO, PMCOMDATE, ADDDATE, ADDUSER, DATE_RAND, RANDBY, SFDATE FROM BIL_SUBJ WHERE (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND STATUS_INFO = 'Randomized' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID   ";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Subject_Visit" + DateTime.Now.ToString("MM-dd-yyyy"));


            worksheet.Cell(1, 1).Value = "Site ID";
            worksheet.Cell(1, 2).Value = "Subject ID";
            worksheet.Cell(1, 3).Value = "Year of Birth";
            worksheet.Cell(1, 4).Value = "Sex";
            worksheet.Cell(1, 5).Value = "Status";
            worksheet.Cell(1, 6).Value = "Randomization Date";
            worksheet.Cell(1, 7).Value = "Treatment";
            worksheet.Cell(1, 8).Value = "Visit 2";
            worksheet.Cell(1, 9).Value = "Visit 4";
            worksheet.Cell(1, 10).Value = "Visit 6";
            worksheet.Cell(1, 11).Value = "Visit 8";

            int i = 2;

            while (rdr.Read())
            {
                worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["SITEID"]) ? rdr["SITEID"].ToString() : null;
                worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["SUBJID"]) ? rdr["SUBJID"].ToString() : null;
                worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["BRTHDTC"]) ? rdr["BRTHDTC"].ToString() : null;
                worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["SEX"]) ? rdr["SEX"].ToString() : null;
                worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["STATUS_INFO"]) ? rdr["STATUS_INFO"].ToString() : null;
                worksheet.Cell(i, 6).Value = !Convert.IsDBNull(rdr["DATE_RAND"]) ? rdr["DATE_RAND"].ToString() : null;
                worksheet.Cell(i, 7).Value = !Convert.IsDBNull(rdr["ARM"]) ? rdr["ARM"].ToString() : null;
                int ROWKEY = (int)rdr["ROW_KEY"];
                GetSubjVisitDet(worksheet.Row(i), ROWKEY);

                i++;

            }
            conn.Close();
            worksheet.Columns().AdjustToContents();
            return View(ImportExceltoDatatable2(workbook));

        }


        public void GetSubjVisitDet(IXLRow row, int rowkey)
        {
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            string sqlState;
            sqlState = "SELECT VISITID, REPLACE(UPPER(CONVERT(varchar, ADDDATE, 106)), ' ', '/') AS DATEASSGN " +
                       "FROM [BIL_VISITS] " +
                       "WHERE [ROW_KEY] = " + rowkey + " ORDER BY [ADDDATE]";


            SqlCommand cmd = new SqlCommand(sqlState, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string visitId = rdr["VISITID"].ToString();
                string dateAssgn = rdr["DATEASSGN"].ToString();

                if (!string.IsNullOrEmpty(visitId))
                {
                    int columnIndex;
                    if (int.TryParse(visitId.Replace("Visit", ""), out columnIndex))
                    {
                        // Assuming columnIndex starts from 7 (Visit 2 is the 7th column in your example)
                        // Update the ClosedXML row with visit details
                        try
                        {
                            if (columnIndex == 2)
                                row.Cell(7 + 1).Value = dateAssgn;
                            if (columnIndex == 4)
                                row.Cell(7 + 2).Value = dateAssgn;
                            if (columnIndex == 6)
                                row.Cell(7 + 3).Value = dateAssgn;
                            if (columnIndex == 8)
                                row.Cell(7 + 4).Value = dateAssgn;
                        }
                        catch (Exception ex)
                        {
                            // Handle the exception or log it for further investigation
                            // For example, you can log the values of columnIndex and visitId to identify the issue
                            Console.WriteLine($"Error updating cell for columnIndex {columnIndex}, visitId {visitId}: {ex.Message}");
                        }
                    }
                    else
                    {
                        // Handle the case where parsing the column index fails
                        // You may want to log a warning or handle this situation accordingly
                    }
                }


            }

            if (rdr != null) rdr.Close();
            if (conn != null) conn.Close();
        }


        public IActionResult SubjectVisitReportExcel()
        {
            //string SITEID = HttpContext.Session.GetString("sesCenter");
            //XLWorkbook wb = new XLWorkbook();
            //connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            //DataTable dt = CreateDataTable("SELECT SITEID AS 'Site ID', SUBJID AS 'Subject ID', VISIT AS 'Visit ID', KitNumber AS 'Kit Number', REPLACE(UPPER(CONVERT(varchar, [ASSIGNMENT_DATE], 106)), ' ', '/') AS 'Dispensed Date', ASSIGNED_BY AS 'Dispensed By', IPLblShipExpiryDate AS 'Expiry Date', IPLblShipLotNo AS 'Lot Number', KitRepled AS 'Replacement For', notdone AS 'Visit not Done', ReaYes AS 'Reason(IP dispensed, visit not done)' FROM BIL_IP_RANGE WHERE SITEID is not NULL AND SUBJID is not null ORDER BY SITEID, SUBJID", connectionString);
            ////DataTable dt = CreateDataTable("SELECT BIL_IP_RANGE.SITEID, BIL_IP_RANGE.KitNumber, BIL_IP_RANGE.KITSTAT, BIL_IP_RANGE.SENDBY, REPLACE(UPPER(CONVERT(varchar, BIL_IP_RANGE.SENDDATE, 106)), ' ', '/') AS 'SEND DATE', BIL_IP_RANGE.RECVDBY, REPLACE(UPPER(CONVERT(varchar, BIL_IP_RANGE.RECVDDATE, 106)), ' ', '/') AS 'RECVD DATE', BIL_IP_RANGE.ASSIGNED, BIL_IP_RANGE.ASSIGNMENT_DATE, BIL_IP_RANGE.ASSIGNED_BY, BIL_IP_RANGE.VISIT, BIL_SUBJ.SUBJID, BIL_IP_RANGE.KITCOMM FROM BIL_SUBJ RIGHT OUTER JOIN BIL_IP_RANGE ON BIL_SUBJ.ROW_KEY = BIL_IP_RANGE.ROW_KEY WHERE ((BIL_IP_RANGE.SITEID = '" + SITEID + "' OR '" + SITEID + "' = '(All)') AND ([SENDBY] IS NOT NULL)) ORDER BY BIL_IP_RANGE.SITEID, KitNumber", connectionString);

            //wb.Worksheets.Add(dt, "Subject Visit Report").Columns().AdjustToContents();
            //using (var stream = new MemoryStream())
            //{
            //    wb.SaveAs(stream);
            //    var content = stream.ToArray();

            //    return File(
            //        content,
            //        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            //        "Subject Visit Report" + DateTime.Now.Year + ".xlsx");
            //}
            string sql = "SELECT ROW_KEY, SITEID, SUBJID, BRTHDTC, SEX, AgeGroup, ICDTC,  STATUS_INFO, PMCOMDATE, ADDDATE, ADDUSER, DATE_RAND, RANDBY, SFDATE FROM BIL_SUBJ WHERE (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND STATUS_INFO = 'Randomized' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY SITEID   ";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Subject_Visit" + DateTime.Now.ToString("MM-dd-yyyy"));


            worksheet.Cell(1, 1).Value = "Site ID";
            worksheet.Cell(1, 2).Value = "Subject ID";
            worksheet.Cell(1, 3).Value = "Year of Birth";
            worksheet.Cell(1, 4).Value = "Sex";
            worksheet.Cell(1, 5).Value = "Status";
            worksheet.Cell(1, 6).Value = "Randomization Date";
            worksheet.Cell(1, 7).Value = "Visit 2";
            worksheet.Cell(1, 8).Value = "Visit 4";
            worksheet.Cell(1, 9).Value = "Visit 6";
            worksheet.Cell(1, 10).Value = "Visit 8";

            int i = 2;

            while (rdr.Read())
            {
                worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["SITEID"]) ? rdr["SITEID"].ToString() : null;
                worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["SUBJID"]) ? rdr["SUBJID"].ToString() : null;
                worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["BRTHDTC"]) ? rdr["BRTHDTC"].ToString() : null;
                worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["SEX"]) ? rdr["SEX"].ToString() : null;
                worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["STATUS_INFO"]) ? rdr["STATUS_INFO"].ToString() : null;
                worksheet.Cell(i, 6).Value = !Convert.IsDBNull(rdr["DATE_RAND"]) ? rdr["DATE_RAND"].ToString() : null;

                int ROWKEY = (int)rdr["ROW_KEY"];
                GetSubjVisitDet(worksheet.Row(i), ROWKEY);

                i++;

            }
            conn.Close();
            worksheet.Columns().AdjustToContents();
            worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 10)).Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
            worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 10)).Style.Font.FontColor = XLColor.White;
            worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 10)).Style.Font.Bold = true;
            worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, 10)).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            //worksheet.SheetView.FreezeRows(1);
            //worksheet.SheetView.FreezeColumns(1);
            //worksheet.SheetView.FreezeColumns(2);
            worksheet.RangeUsed().SetAutoFilter();

            using (var stream = new MemoryStream())
            {
                workbook.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   "Subject Visit Report" + DateTime.Now.Year + ".xlsx");
            }



        }


        public IActionResult SiteIPReport()
        {
            string sql = "SELECT SITEID, KitNumber, KITTYPE, SENDBY, RECVDBY, RECVDDATE, KITSTAT, ASSIGNED, VISIT, SUBJID, KITCOMM, ASSIGNMENT_DATE, SENDDATE, IPLblShipExpiryDate, IPLblShipLotNo, ASSIGNED_BY, KitRepled FROM BIL_IP_RANGE WHERE (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND (SENDBY IS NOT NULL) ORDER BY SITEID, KitNumber";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("SiteIP" + DateTime.Now.ToString("MM-dd-yyyy"));


            worksheet.Cell(1, 1).Value = "Site ID";
            worksheet.Cell(1, 2).Value = "Kit Number";
            worksheet.Cell(1, 3).Value = "Kit Type";
            worksheet.Cell(1, 4).Value = "Kit Status";
            worksheet.Cell(1, 5).Value = "Received By";
            worksheet.Cell(1, 6).Value = "Received Date";
            worksheet.Cell(1, 7).Value = "Sent By";
            worksheet.Cell(1, 8).Value = "Sent Date";
            worksheet.Cell(1, 9).Value = "Assigned";
            worksheet.Cell(1, 10).Value = "Subject ID";
            worksheet.Cell(1, 11).Value = "Visit ID";
            worksheet.Cell(1, 12).Value = "Assigment Date";
            worksheet.Cell(1, 13).Value = "Assigned By";
            worksheet.Cell(1, 14).Value = "Expiry Date";
            worksheet.Cell(1, 15).Value = "Lot NUmber";
            worksheet.Cell(1, 16).Value = "Replacement For";
            worksheet.Cell(1, 17).Value = "Comments";


            int i = 2;

            while (rdr.Read())
            {
                worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["SITEID"]) ? rdr["SITEID"].ToString() : null;
                worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["KitNumber"]) ? rdr["KitNumber"].ToString() : null;
                worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["KITTYPE"]) ? rdr["KITTYPE"].ToString() : null;
                worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["KITSTAT"]) ? rdr["KITSTAT"].ToString() : null;
                worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["RECVDBY"]) ? rdr["RECVDBY"].ToString() : null;
                
                DateTime? recvDate = !Convert.IsDBNull(rdr["RECVDDATE"]) ? Convert.ToDateTime(rdr["RECVDDATE"]) : (DateTime?)null;
                string formattedDate = recvDate.HasValue ? recvDate.Value.ToString("dd/MMM/yyyy") : null;
                worksheet.Cell(i, 6).Value = formattedDate;
                worksheet.Cell(i, 7).Value = !Convert.IsDBNull(rdr["SENDBY"]) ? rdr["SENDBY"].ToString() : null;
                worksheet.Cell(i, 8).Value = !Convert.IsDBNull(rdr["SENDDATE"]) ? rdr["SENDDATE"].ToString() : null;
                worksheet.Cell(i, 9).Value = !Convert.IsDBNull(rdr["ASSIGNED"]) ? rdr["ASSIGNED"].ToString() : null;
                worksheet.Cell(i, 10).Value = !Convert.IsDBNull(rdr["SUBJID"]) ? rdr["SUBJID"].ToString() : null;
                worksheet.Cell(i, 11).Value = !Convert.IsDBNull(rdr["VISIT"]) ? rdr["VISIT"].ToString() : null;
                worksheet.Cell(i, 12).Value = !Convert.IsDBNull(rdr["ASSIGNMENT_DATE"]) ? rdr["ASSIGNMENT_DATE"].ToString() : null;
                worksheet.Cell(i, 13).Value = !Convert.IsDBNull(rdr["ASSIGNED_BY"]) ? rdr["ASSIGNED_BY"].ToString() : null;
                worksheet.Cell(i, 14).Value = !Convert.IsDBNull(rdr["IPLblShipExpiryDate"]) ? rdr["IPLblShipExpiryDate"].ToString() : null;
                worksheet.Cell(i, 15).Value = !Convert.IsDBNull(rdr["IPLblShipLotNo"]) ? rdr["IPLblShipLotNo"].ToString() : null;
                worksheet.Cell(i, 16).Value = !Convert.IsDBNull(rdr["KitRepled"]) ? rdr["KitRepled"].ToString() : null;
                worksheet.Cell(i, 17).Value = !Convert.IsDBNull(rdr["KITCOMM"]) ? rdr["KITCOMM"].ToString() : null;
                //worksheet.Cell(i, 10).Value = !Convert.IsDBNull(rdr["notdone"]) ? rdr["notdone"].ToString() : null;
                //worksheet.Cell(i, 11).Value = !Convert.IsDBNull(rdr["ReaYes"]) ? rdr["ReaYes"].ToString() : null;
                //worksheet.Cell(i, 12).Value = !Convert.IsDBNull(rdr["SFDATE"]) ? rdr["SFDATE"].ToString() : null;
                i++;

            }
            conn.Close();
            worksheet.Columns().AdjustToContents();
            return View(ImportExceltoDatatable(workbook));

        }

        public IActionResult SiteIPReportExcel()
        {

            XLWorkbook wb = new XLWorkbook();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DataTable dt = CreateDataTable("SELECT SITEID AS 'Site ID', KitNumber AS 'Kit Number', KITTYPE AS 'Kit Type', KITSTAT AS 'Kit Status',  RECVDBY AS 'Received By', RECVDDATE AS 'Received Date', SENDBY AS 'Send By', SENDDATE AS 'Sent Date', ASSIGNED AS 'Assigned', SUBJID AS 'Subject ID', VISIT AS 'Visit ID', ASSIGNMENT_DATE AS 'Dispensed Date',  ASSIGNED_BY AS 'Assigned By', IPLblShipExpiryDate AS 'Expiry Date', IPLblShipLotNo AS 'Lot Number', KitRepled AS 'Replacement For' , KITCOMM AS 'Comments' FROM BIL_IP_RANGE WHERE (SITEID = '" + HttpContext.Session.GetString("sesCenter") + "' OR '" + HttpContext.Session.GetString("sesCenter") + "' = '(All)') AND (SENDBY IS NOT NULL)  ORDER BY SITEID, KitNumber", connectionString);
            wb.Worksheets.Add(dt, "Site IP Report").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Site IP Report" + DateTime.Now.Year + ".xlsx");
            }

        }

        public IActionResult DepotIPReport()
        {
            string sql = "SELECT KITKEY, KitNumber, KITSTAT, IPLblShipExpiryDate, IPLblShipLotNo, KITCOMM FROM BIL_IP_RANGE WHERE (SITEID IS  NULL)  ORDER BY  KitNumber";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader rdr = cmd.ExecuteReader();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("IPDepot" + DateTime.Now.ToString("MM-dd-yyyy"));


            worksheet.Cell(1, 1).Value = "Kit Key";
            worksheet.Cell(1, 2).Value = "Kit Number";
            worksheet.Cell(1, 3).Value = "Kit Status";
            worksheet.Cell(1, 4).Value = "Expiry Date";
            worksheet.Cell(1, 5).Value = "Lot NUmber";
            worksheet.Cell(1, 6).Value = "Comments";


            int i = 2;

            while (rdr.Read())
            {
                worksheet.Cell(i, 1).Value = !Convert.IsDBNull(rdr["KITKEY"]) ? rdr["KITKEY"].ToString() : null;
                worksheet.Cell(i, 2).Value = !Convert.IsDBNull(rdr["KitNumber"]) ? rdr["KitNumber"].ToString() : null;
                worksheet.Cell(i, 3).Value = !Convert.IsDBNull(rdr["KITSTAT"]) ? rdr["KITSTAT"].ToString() : null;
                worksheet.Cell(i, 4).Value = !Convert.IsDBNull(rdr["IPLblShipExpiryDate"]) ? rdr["IPLblShipExpiryDate"].ToString() : null;
                worksheet.Cell(i, 5).Value = !Convert.IsDBNull(rdr["IPLblShipLotNo"]) ? rdr["IPLblShipLotNo"].ToString() : null;
                worksheet.Cell(i, 6).Value = !Convert.IsDBNull(rdr["KITCOMM"]) ? rdr["KITCOMM"].ToString() : null;
                //worksheet.Cell(i, 10).Value = !Convert.IsDBNull(rdr["notdone"]) ? rdr["notdone"].ToString() : null;
                //worksheet.Cell(i, 11).Value = !Convert.IsDBNull(rdr["ReaYes"]) ? rdr["ReaYes"].ToString() : null;
                //worksheet.Cell(i, 12).Value = !Convert.IsDBNull(rdr["SFDATE"]) ? rdr["SFDATE"].ToString() : null;
                i++;

            }
            conn.Close();
            worksheet.Columns().AdjustToContents();
            return View(ImportExceltoDatatable(workbook));

        }

        public IActionResult DepotIPReportExcel()
        {

            XLWorkbook wb = new XLWorkbook();
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            DataTable dt = CreateDataTable("SELECT KITKEY AS 'Kit Key', KitNumber AS 'Kit Number', KITSTAT AS 'Kit Status',  IPLblShipExpiryDate AS 'Expiry Date', IPLblShipLotNo AS 'Lot Number', KITCOMM AS 'Comments' FROM BIL_IP_RANGE WHERE (SITEID IS  NULL) ORDER BY  KitNumber", connectionString);
            wb.Worksheets.Add(dt, "IPDepot Report").Columns().AdjustToContents();
            using (var stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                var content = stream.ToArray();

                return File(
                    content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "IP Depot Report" + DateTime.Now.Year + ".xlsx");
            }

        }

        private DataTable CreateDataTable(string cmdText, string connection)
        {
            System.Data.DataTable dt = new DataTable();
            System.Data.SqlClient.SqlDataAdapter da = new SqlDataAdapter(cmdText, connection);
            da.Fill(dt);
            return dt;
        }
    }
}
