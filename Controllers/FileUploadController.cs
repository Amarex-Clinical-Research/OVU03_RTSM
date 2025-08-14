using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using Webview_IRT.Models;
using Spire.Xls;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace RTSM_OLSingleArm.Controllers
{
    public class FileUploadController : Controller
    {
        public string connectionString;
        readonly IConfiguration _configuration;
        public FileUploadController(IConfiguration configuration)
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
        public IActionResult FileUploadHome()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            String sql = "SELECT UPLOAD_KEY, UPLOAD_DESC, FILENAME, EMAILSENT, TYPEUL, ADDUSER, ADDDATE, EmailAddress, CodeSend FROM IP_BIO_UPLOADS WHERE TYPEUL = 'Kit' AND IsHide is null AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY  FILENAME, UPLOAD_KEY , ADDDATE ";
            String sql2 = "SELECT Upload_KEY, UPLOAD_DESC, FILENAME, EMAILSENT, TYPEUL, ADDUSER, ADDDATE, EmailAddress, CodeSend FROM IP_BIO_UPLOADS WHERE TYPEUL = 'Rand' AND IsHide is null AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY  FILENAME, UPLOAD_KEY , ADDDATE ";
            String sql3 = "SELECT Upload_KEY, UPLOAD_DESC, FILENAME, EMAILSENT, TYPEUL, ADDUSER, ADDDATE, CHANGEUSER, CHANGEDATE, Reason, EmailAddress, CodeSend FROM IP_BIO_UPLOADS WHERE IsHide = 'Yes' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ORDER BY  FILENAME, UPLOAD_KEY , ADDDATE ";

            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlCommand cmd2 = new SqlCommand(sql2, con);
            SqlCommand cmd3 = new SqlCommand(sql3, con);

            var list = new IPBioList();
            var kit = new List<KitUpload>();
            var rand = new List<RandUpload>();
            var remove = new List<RemovedKits>();


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
                    temp.EmailAddress = rdr["EmailAddress"].ToString();
                    temp.CodeSend = rdr["CodeSend"].ToString();
                    kit.Add(temp);
                }
                rdr.Close();
                SqlDataReader rdr3 = cmd2.ExecuteReader();
                while (rdr3.Read())
                {
                    var temp = new RandUpload();
                    temp.UPLOAD_KEY = (int)rdr3["UPLOAD_KEY"];
                    temp.UPLOAD_DESC = rdr3["UPLOAD_DESC"].ToString();
                    temp.FILENAME = rdr3["FILENAME"].ToString();
                    temp.EMAILSENT = rdr3["EMAILSENT"].ToString();
                    temp.TYPEUL = rdr3["TYPEUL"].ToString();
                    temp.ADDUSER = rdr3["ADDUSER"].ToString();
                    temp.ADDDATE = (DateTime)rdr3["ADDDATE"];
                    temp.EmailAddress = rdr3["EmailAddress"].ToString();
                    temp.CodeSend = rdr3["CodeSend"].ToString();
                    rand.Add(temp);
                }
                rdr3.Close();
                SqlDataReader rdr4 = cmd3.ExecuteReader();
                while (rdr4.Read())
                {
                    var temp = new RemovedKits();
                    temp.UPLOAD_KEY = (int)rdr4["UPLOAD_KEY"];
                    temp.UPLOAD_DESC = rdr4["UPLOAD_DESC"].ToString();
                    temp.FILENAME = rdr4["FILENAME"].ToString();
                    temp.EMAILSENT = rdr4["EMAILSENT"].ToString();
                    temp.TYPEUL = rdr4["TYPEUL"].ToString();
                    temp.ADDUSER = rdr4["ADDUSER"].ToString();
                    temp.ADDDATE = (DateTime)rdr4["ADDDATE"];
                    temp.CHANGEUSER = rdr4["CHANGEUSER"].ToString();
                    temp.CHANGEDATE = (DateTime)rdr4["CHANGEDATE"];
                    temp.Reason = rdr4["Reason"].ToString();
                    temp.EmailAddress = rdr4["EmailAddress"].ToString();
                    temp.CodeSend = rdr4["CodeSend"].ToString();
                    remove.Add(temp);
                }
                rdr4.Close();

            }
            con.Close();
            list.kitupload = kit;
            list.randupload =rand;
            list.removed = remove;
            string emails = GetEmailList();
            TempData["email"] = emails;
            return View(list);
           
        }
      

        public IActionResult UploadKit(IFormFile File, string comment, string email, bool Enable)
        {
            Enable = true;
            //bool value = CheckExcelFilePassword(File);
            try
            {
                string[] allowedExtensions = { ".xls", ".xlsx" };
                var fileExtension = Path.GetExtension(File.FileName).ToLower();
               email = email.Replace(",", ";");
                string[] emails = email.Split(';');

                foreach (string emailAddress in emails)
                {
                    // Trim any leading or trailing spaces from the email address
                    string trimmedEmail = emailAddress.Trim();

                    // Check if the trimmed email matches a valid email pattern
                    if (!IsValidEmailAddress(trimmedEmail))
                    {
                        TempData["ErrorMessage"] = "Invalid email address format: " + trimmedEmail;
                        return RedirectToAction("FileUploadHome");
                    }
                }
                if (allowedExtensions.Contains(fileExtension))
                {

                    bool value = CheckExcelFilePassword(File);
                    if (!value)
                    {
                        TempData["ErrorMessage"] = "The uploaded file does not have password protection. Kindly upload only the password-protected file to proceed.";
                        return RedirectToAction("FileUploadHome");
                    }
                    Byte[] bytes = null;

                    using (MemoryStream ms = new MemoryStream())
                    {
                        File.OpenReadStream().CopyTo(ms);
                        bytes = ms.ToArray();
                    }

                    connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                    SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
                    if (Enable)
                    {
                        SqlCommand cmd = new SqlCommand("insert into IP_BIO_UPLOADS (SPKEY, FILENAME, TYPEUL, UplData,  ADDUSER, ADDDATE, EMAILSENT, EmailAddress, CodeSend) values (@SPKEY, @FILENAME, @TYPEUL, @UplData ,@ADDUSER, SYSDATETIME(), @EMAILSENT, @EmailAddress, @CodeSend)", con);
                        cmd.CommandType = CommandType.Text;

                        cmd.Parameters.AddWithValue("@SPKEY", int.Parse(HttpContext.Session.GetString("sesSPKey")));
                        cmd.Parameters.AddWithValue("@FILENAME", File.FileName);
                        cmd.Parameters.AddWithValue("@UplData", bytes);
                        cmd.Parameters.AddWithValue("@TYPEUL", "Kit");
                        cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));
                        cmd.Parameters.AddWithValue("@EMAILSENT", "Yes");
                        cmd.Parameters.AddWithValue("@EmailAddress", email);
                        cmd.Parameters.AddWithValue("@CodeSend", "No");

                        con.Open();
                        ViewBag.ID = cmd.ExecuteScalar();
                        con.Close();


                        string UserName = GetUserName(HttpContext.Session.GetString("suserid"));
                        string message = "Protocol: " + StudyName() + "\n";
                        message += "Please find the attached kit list document. For security purpose, this document is password-protected, and you'll receive a follow-up email shortly with the password. If not received within an hour, or for any concerns, please submit a problem report." + "\n" + "\n";
                        message += "Kit List Released By: " + UserName + "\n" + "FileName: " + File.FileName + "\n"+ "\n" + "Addtional Details: " + comment + "\n";
                        string subject = StudyName() + " - WebView RTSM - Master kit list upload notification";
                        email = email + ";" + GetEmail(HttpContext.Session.GetString("suserid"));
                        SendEmailWithAttachment(email, subject, message, bytes, File.FileName);
                        
                        TempData["Message"] = "File sent successfully. Kindly utilize the designated action button to provide the password for the uploaded file.";
                    }
                    else
                    {
                        SqlCommand cmd = new SqlCommand("insert into IP_BIO_UPLOADS (SPKEY, FILENAME, TYPEUL, UplData,  ADDUSER, ADDDATE) values (@SPKEY, @FILENAME, @TYPEUL, @UplData ,@ADDUSER, SYSDATETIME())", con);
                        cmd.CommandType = CommandType.Text;

                        cmd.Parameters.AddWithValue("@SPKEY", int.Parse(HttpContext.Session.GetString("sesSPKey")));
                        cmd.Parameters.AddWithValue("@FILENAME", File.FileName);
                        cmd.Parameters.AddWithValue("@UplData", bytes);
                        cmd.Parameters.AddWithValue("@TYPEUL", "Kit");
                        cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));
          
                        con.Open();
                        ViewBag.ID = cmd.ExecuteScalar();
                        con.Close();
                        TempData["Message"] = "File uploaded";
                    }
                }
                else
                {
                    TempData["ErrorMessage"] = "Invalid file type. Only excel files(.xls, .xlsx) are permitted. Please ensure that you are uploading the file with the correct format.";
                }
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = "An error occurred while uploading the file: " + ex.Message;
            }

            return RedirectToAction("FileUploadHome");
        }


        public IActionResult UploadRand(IFormFile File, string comment, string email, bool Enable) 
        {
            Enable = true;
            try

            {
                string[] allowedExtensions = { ".xls", ".xlsx" };
                var fileExtension = Path.GetExtension(File.FileName).ToLower();

                email = email.Replace(",", ";");
                string[] emails = email.Split(';');

                foreach (string emailAddress in emails)
                {
                    // Trim any leading or trailing spaces from the email address
                    string trimmedEmail = emailAddress.Trim();

                    // Check if the trimmed email matches a valid email pattern
                    if (!IsValidEmailAddress(trimmedEmail))
                    {
                        TempData["ErrorMessage"] = "Invalid email address format: " + trimmedEmail;
                        return RedirectToAction("FileUploadHome");
                    }
                }
                if (allowedExtensions.Contains(fileExtension))
                {
                    bool value = CheckExcelFilePassword(File);
                    if (!value)
                    {
                        TempData["ErrorMessage"] = "The uploaded file does not have password protection. Kindly upload only the password-protected file to proceed.";
                        return RedirectToAction("FileUploadHome");
                    }

                    Byte[] bytes = null;
                    using (MemoryStream ms = new MemoryStream())

                    {

                        File.OpenReadStream().CopyTo(ms);

                        bytes = ms.ToArray();

                    }
                        connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
                        SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);

                    if (Enable)
                    {
                        SqlCommand cmd = new SqlCommand("insert into IP_BIO_UPLOADS (SPKEY, FILENAME, TYPEUL, UplData,  ADDUSER, ADDDATE, EMAILSENT, EmailAddress, CodeSend) values (@SPKEY, @FILENAME, @TYPEUL, @UplData ,@ADDUSER, SYSDATETIME(), @EMAILSENT, @EmailAddress, @CodeSend)", con);
                        cmd.CommandType = CommandType.Text;

                        cmd.Parameters.AddWithValue("@SPKEY", int.Parse(HttpContext.Session.GetString("sesSPKey")));

                        cmd.Parameters.AddWithValue("@FILENAME", File.FileName);

                        cmd.Parameters.AddWithValue("@UplData", bytes);

                        cmd.Parameters.AddWithValue("@TYPEUL", "Rand");

                        cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));

                        cmd.Parameters.AddWithValue("@EMAILSENT", "Yes");
                        cmd.Parameters.AddWithValue("@EmailAddress", email);
                        cmd.Parameters.AddWithValue("@CodeSend", "No");



                        con.Open();

                        ViewBag.ID = cmd.ExecuteScalar();
                        ViewBag.FileName = File.FileName;

                        con.Close();
                        string UserName = GetUserName(HttpContext.Session.GetString("suserid"));
                        string message = "Protocol: " + StudyName() + "\n";
                        message += "Please find the attached kit list document. For security purpose, this document is password-protected, and you'll receive a follow-up email shortly with the password. If not received within an hour, or for any concerns, please submit a problem report." + "\n" + "\n";
                        message += "Randomization List Released By: " + UserName + "\n" + "FileName: " + File.FileName + "\n" + "\n" + "Addtional Details: " + comment + "\n";
                        string subject = StudyName() + " - WebView RTSM - Subject Randomization List upload notification";
                        email = email + ";" + GetEmail(HttpContext.Session.GetString("suserid"));
                        SendEmailWithAttachment(email, subject, message, bytes, File.FileName);

                        TempData["Message"] = "File sent successfully. Kindly utilize the designated action button to provide the password for the uploaded file.";
                    }
                    else
                    {
                        SqlCommand cmd = new SqlCommand("insert into IP_BIO_UPLOADS (SPKEY, FILENAME, TYPEUL, UplData,  ADDUSER, ADDDATE) values (@SPKEY, @FILENAME, @TYPEUL, @UplData ,@ADDUSER, SYSDATETIME())", con);
                        cmd.CommandType = CommandType.Text;

                        cmd.Parameters.AddWithValue("@SPKEY", int.Parse(HttpContext.Session.GetString("sesSPKey")));
                        cmd.Parameters.AddWithValue("@FILENAME", File.FileName);
                        cmd.Parameters.AddWithValue("@UplData", bytes);
                        cmd.Parameters.AddWithValue("@TYPEUL", "Rand");
                        cmd.Parameters.AddWithValue("@ADDUSER", HttpContext.Session.GetString("suserid"));

                        con.Open();
                        ViewBag.ID = cmd.ExecuteScalar();
                        con.Close();
                        TempData["Message"] = "File uploaded";
                      

                    }

                }

                else
                {
                    TempData["ErrorMessage"] = "Invalid file type. Only excel files(.xls, .xlsx) are permitted. Please ensure that you are uploading the file with the correct format.";
                }
            }
            catch (Exception ex)

            {
                TempData["ErrorMessage"] = "An error occurred while uploading the file: " + ex.Message;
            }

            return RedirectToAction("FileUploadHome");
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

                        //filetype = "application/pdf";
                        //filetype = "Application/x-ms excel";
                        filetype = "application/xlsx";

                        name = Convert.ToString(sdr["FILENAME"]);

                    }

                    con.Close();

                }

            }

            return File(bytes, filetype, name, lastModified: DateTime.UtcNow.AddSeconds(-5),

        entityTag: new Microsoft.Net.Http.Headers.EntityTagHeaderValue("\"MyCalculatedEtagValue\""));

        }

        //public ActionResult GetEmailList()
        //{
        //    int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
        //    string userid = HttpContext.Session.GetString("suserid");
        //    string SITEID = HttpContext.Session.GetString("sesCenter");
        //    connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

        //   String biosql = "SELECT KEYID, PIDet, PIType  FROM Email_Notifications WHERE PIType = 'PMIPLblKitListUL' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";

        //    SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);            
        //    SqlCommand cmd = new SqlCommand(biosql, con);
        //    BiometircUploadEmailNoti model = new BiometircUploadEmailNoti();
        //    con.Open();
        //    SqlDataReader rdr = cmd.ExecuteReader();
        //    List<string> valuesList = new List<string>(); // Initialize a list to store trimmed values

        //    while (rdr.Read())
        //    {
        //        string value = rdr["PIDet"].ToString().Trim(); // Get and trim the value of PIDet column
        //        if (!string.IsNullOrEmpty(value))
        //        {
        //            valuesList.Add(value); // Add the trimmed value to the list
        //        }
        //    }

        //    string concatenatedValues = string.Join(",", valuesList); // Join the trimmed values using semicolons
        //    model.PIDet = concatenatedValues; // Assign the concatenated string to the model.PIDet property
        //    TempData["email"] = concatenatedValues;
        //    rdr.Close();
        //    con.Close();

        //    //return View(list);

        //    return PartialView("_GetEmailList", model);

        //}

        public string GetEmailList()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

            String biosql = "SELECT KEYID, PIDet, PIType  FROM Email_Notifications WHERE PIType = 'PMIPLblKitListUL' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";

            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(biosql, con);
            BiometircUploadEmailNoti model = new BiometircUploadEmailNoti();
            con.Open();
            SqlDataReader rdr = cmd.ExecuteReader();
            List<string> valuesList = new List<string>(); // Initialize a list to store trimmed values

            while (rdr.Read())
            {
                string value = rdr["PIDet"].ToString().Trim(); // Get and trim the value of PIDet column
                if (!string.IsNullOrEmpty(value))
                {
                    valuesList.Add(value); // Add the trimmed value to the list
                }
            }

            string concatenatedValues = string.Join(";", valuesList); // Join the trimmed values using semicolons
            model.PIDet = concatenatedValues; // Assign the concatenated string to the model.PIDet property
            TempData["email"] = concatenatedValues;
            rdr.Close();
            con.Close();

            //return View(list);

            return concatenatedValues;

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

        public string GetUserName(string userid)
        {
            string name = "";
            string email = "";
            string sql = "SELECT User_Email, LName, FName FROM zSecurityID WHERE (UserID = '" + userid + " ')";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("AmarexDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    email = rdr["User_Email"].ToString();
                    name = rdr["FName"].ToString();
                    name += " " + rdr["LName"].ToString();

                } 
            }
            return name;
        }

        public ActionResult RandModel()
        {
            int SPKEY = int.Parse(HttpContext.Session.GetString("sesSPKey"));
            string userid = HttpContext.Session.GetString("suserid");
            string SITEID = HttpContext.Session.GetString("sesCenter");
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");

            String biosql = "SELECT KEYID, PIDet, PIType  FROM Email_Notifications WHERE PIType = 'PMIPLblKitListUL' AND SPKEY = " + HttpContext.Session.GetString("sesSPKey") + " ";

            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(biosql, con);
            BiometircUploadEmailNoti model = new BiometircUploadEmailNoti();
            con.Open();
            SqlDataReader rdr = cmd.ExecuteReader();
            List<string> valuesList = new List<string>(); // Initialize a list to store trimmed values

            while (rdr.Read())
            {
                string value = rdr["PIDet"].ToString().Trim(); // Get and trim the value of PIDet column
                if (!string.IsNullOrEmpty(value))
                {
                    valuesList.Add(value); // Add the trimmed value to the list
                }
            }

            string concatenatedValues = string.Join(",", valuesList); // Join the trimmed values using semicolons
            model.PIDet = concatenatedValues; // Assign the concatenated string to the model.PIDet property
            TempData["email"] = concatenatedValues;
            rdr.Close();
            con.Close();

            //return View(list);

            return PartialView("_UploadKit", model);

        }

        public IActionResult RemoveKit(int UPLOAD_KEY, string FILENAME)
        {
            string sql = "delete from IP_BIO_UPLOADS where UPLOAD_KEY = @UPLOAD_KEY";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            conn.Open();
            SqlCommand cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@UPLOAD_KEY", UPLOAD_KEY);
            cmd.ExecuteNonQuery();
            conn.Close();
            return RedirectToAction("FileUploadHome");
        }

        [HttpGet]
        public ActionResult GetUploadKey(int UPLOAD_KEY)
        {
            BiometircUploadEmailNoti model = new BiometircUploadEmailNoti();
            connectionString = _configuration.GetConnectionString("IRTDB");
            string sql2 = "SELECT * FROM IP_BIO_UPLOADS WHERE UPLOAD_KEY = @UPLOAD_KEY";
            SqlConnection con = new SqlConnection(connectionString);
            using (con)
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(sql2, con);
                cmd.Parameters.AddWithValue("@UPLOAD_KEY", UPLOAD_KEY);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    
                    ViewBag.FILENAME = rdr["FILENAME"].ToString();
                }

            }

            model.UPLOAD_KEY = UPLOAD_KEY;
            return PartialView("_AddReason", model);

        }

        public ActionResult HideKit(int UPLOAD_KEY, string Reason, string FILENAME)
        {
            string userid = HttpContext.Session.GetString("suserid");
            string sql = "Update IP_BIO_UPLOADS SET IsHide = @ISHide, Reason = @Reason, CHANGEUSER = @CHANGEUSER, CHANGEDATE =SYSDATETIME() WHERE UPLOAD_KEY = @UPLOAD_KEY";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();

                cmd.Parameters.AddWithValue("@UPLOAD_KEY", UPLOAD_KEY);
                cmd.Parameters.AddWithValue("@IsHide", "Yes");
                cmd.Parameters.AddWithValue("@CHANGEUSER", userid);
                cmd.Parameters.AddWithValue("@Reason", Reason);

                cmd.ExecuteNonQuery();

            }
            conn.Close();
            TempData["Message"] = "File: " + FILENAME +" removed from the table";
            return RedirectToAction("FileUploadHome");

        }

        public IActionResult RestoreFile(int UPLOAD_KEY, string FILENAME)
        {
            string userid = HttpContext.Session.GetString("suserid");
            string sql = "Update IP_BIO_UPLOADS SET IsHide = null, Reason = null, CHANGEUSER = @CHANGEUSER, CHANGEDATE =SYSDATETIME() WHERE UPLOAD_KEY = @UPLOAD_KEY";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();

                cmd.Parameters.AddWithValue("@UPLOAD_KEY", UPLOAD_KEY);
                cmd.Parameters.AddWithValue("@CHANGEUSER", userid);
                

                cmd.ExecuteNonQuery();

            }
            conn.Close();
            TempData["Message"] = "File: " +FILENAME + " restored" ;
            return RedirectToAction("FileUploadHome");
        }

        public string GetEmail(string userID)
        {
            string sql = "SELECT zSecurityID.User_Email FROM zSecurityID WHERE (UserID = '" + userID + "') ";
            connectionString = _configuration.GetConnectionString("AmarexDbConnStr");
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {

                    object result = cmd.ExecuteScalar();
                    return result.ToString();
                }
            }


        }

        private bool CheckExcelFilePassword(IFormFile file)
        {
            try
            {
                
                var workbook = new Workbook();
                workbook.LoadFromStream(file.OpenReadStream());
         
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Invalid password"))
                {
                    return true; 
                }
                return false;
            }

            return false; // If no exception is thrown, the file is not password-protected
        }


        //public void SendEmailWithAttachment(string SendTo, string Subject, string message, byte[] attachmentData, string attachmentFileName)
        //{
        //    connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
        //    SqlConnection conn = new SqlConnection(connectionString);

        //    SqlCommand EmailCmd = new SqlCommand("NotifySendwithAttachment2", conn);
        //    EmailCmd.CommandType = CommandType.StoredProcedure;
        //    EmailCmd.Parameters.Add("@SENDMAILTO", SqlDbType.VarChar).Value = SendTo;
        //    EmailCmd.Parameters.Add("@SUBJ", SqlDbType.VarChar).Value = Subject;
        //    EmailCmd.Parameters.Add("@MSGBODY", SqlDbType.VarChar).Value = message;

        //    // Add parameters for attachment
        //    EmailCmd.Parameters.Add("@ATTACHMENTDATA", SqlDbType.VarBinary).Value = attachmentData;
        //    EmailCmd.Parameters.Add("@ATTACHMENTFILENAME", SqlDbType.VarChar).Value = attachmentFileName;

        //    using (conn)
        //    {
        //        conn.Open();
        //        EmailCmd.ExecuteReader();
        //    }
        //}

        public void SendEmailWithAttachment(string toEmail, string subject, string body, byte[] attachmentData, string attachmentFileName)
        {
            try
            {
                string[] notificationEmail = toEmail.Split(';');
                SmtpClient sc = new SmtpClient("");
                sc.Host = "192.168.154.30";
                sc.UseDefaultCredentials = false;
                sc.Credentials = new NetworkCredential("CRO\\donot-reply", "wiev246*");
                sc.EnableSsl = false;
                //var client = new SmtpClient("your-smtp-server.com");
                string bodyText = "--------------------------------------------------------------------------------------------------------------------------"
                             + "\n\n"
                             + "This email message, from Amarex LLC, may contain confidential information, intended only for the designated recipient(s). If you are not the designated recipient, you are hereby notified that any disclosure, " +
                                "copying, distribution, use of, or reliance on, the contents of this e - mail is prohibited.If this message was received in error, please contact the sender by reply email and destroy all copies of the original message.";
                MailMessage msg = new MailMessage();
                msg.From = new MailAddress("donot-reply@amarexcro.com", "Amarex donot-reply");

                msg.Subject = subject;
                msg.Body = body + "\n\n" + bodyText;
                foreach (string email in notificationEmail)
                {
                    string trimmedEmail = email.Trim(';');
                    if (!string.IsNullOrEmpty(trimmedEmail))
                    {
                        if (IsValidEmail(trimmedEmail))
                        {
                            msg.To.Add(new MailAddress(trimmedEmail));
                        }

                    }
                }


                // Attach the file to the email
                var attachment = new Attachment(new MemoryStream(attachmentData), attachmentFileName);
                        msg.Attachments.Add(attachment);

                        // Send the email
                        sc.Send(msg);
                    
                

                // Additional logic after sending the email, if needed
            }
            catch (Exception ex)
            {
                // Handle email sending exceptions
                TempData["ErrorMessage"] = "An error occurred while sending the email: " + ex.Message;
            }
        }

        bool IsValidEmail(string email)
        {
            var trimmedEmail = email.Trim();

            if (trimmedEmail.EndsWith("."))
            {
                return false; // suggested by @TK-421
            }
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == trimmedEmail;
            }
            catch
            {
                return false;
            }
        }


        public IActionResult SendCode(string code, string comment, string email, int UPLOAD_KEY, string FILENAME, string TYPEUL)
        {

            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            SqlConnection con = new System.Data.SqlClient.SqlConnection(connectionString);
            email = email.Replace(",", ";");
            string[] emails = email.Split(';');

            foreach (string emailAddress in emails)
            {
                // Trim any leading or trailing spaces from the email address
                string trimmedEmail = emailAddress.Trim();

                // Check if the trimmed email matches a valid email pattern
                if (!IsValidEmailAddress(trimmedEmail))
                {
                    TempData["ErrorMessage"] = "Invalid email address format: " + trimmedEmail;
                    return RedirectToAction("FileUploadHome");
                }
            }
            bool rtnVal = CheckEmailAddress(UPLOAD_KEY, email);
            if (!rtnVal)
            {
                TempData["ErrorMessage"] = "Email address does not match.";
                return RedirectToAction("FileUploadHome");
            }

            string sql = "Update IP_BIO_UPLOADS SET CodeSend = 'Yes', SendDate =SYSDATETIME() WHERE UPLOAD_KEY = @UPLOAD_KEY";
            SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("VpeRandDbConnStr"));
            SqlCommand cmd = new SqlCommand(sql, conn);
            using (conn)
            {
                conn.Open();
                cmd.Parameters.AddWithValue("@UPLOAD_KEY", UPLOAD_KEY);
                cmd.ExecuteNonQuery();

            }
            string UserName = GetUserName(HttpContext.Session.GetString("suserid"));
            if (TYPEUL == "Kit")
            {
                TYPEUL = "Kit";
            }
            if (TYPEUL == "Rand")
            {
                TYPEUL = "Subject randomization";
            }
            string message = "Protocol: " + StudyName() + "\n" + "\n";
            message += TYPEUL + " list Released By: " + UserName + "\n" + "FileName: " + FILENAME + "\n" + "\n" + "Password: "+ code +"\n" + "Addtional Details: " + comment + "\n";
            string subject = StudyName() + " - WebView RTSM - Master " + TYPEUL + " list password notification";
            email = email + ";" + GetEmail(HttpContext.Session.GetString("suserid"));
            SendEmail(email, subject, message);
            TempData["Message"] = "Password sent for File: " + "'"+FILENAME+"'";
            return RedirectToAction("FileUploadHome");
        }

        public bool CheckEmailAddress(int Uploadkey, string email)
        {
            string rtnVal = "";
            connectionString = _configuration.GetConnectionString("VpeRandDbConnStr");
            string sql2 = "SELECT EmailAddress FROM IP_BIO_UPLOADS WHERE  UPLOAD_KEY = " + Uploadkey + "";
            SqlConnection con = new SqlConnection(connectionString);
            using (con)
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(sql2, con);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    rtnVal = rdr["EmailAddress"].ToString();
                    
                }

            }
            if(rtnVal != "")
            {
                string[] emails = email.Split(';');
                foreach (string emailAddress in emails)
                {
                    if (rtnVal.Contains(emailAddress))
                    {
                        
                    }
                    else
                    {
                        return false;
                    }

                }
            }
            else
            {
                return false;
            }
            return true;
        }

        private bool IsValidEmailAddress(string email)
        {
            // Define a regex pattern for email validation
            string pattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";

            // Create a Regex object with the pattern
            Regex regex = new Regex(pattern);

            // Use the Regex object's Match method to check if the email matches the pattern
            return regex.IsMatch(email);
        }

    }
}
