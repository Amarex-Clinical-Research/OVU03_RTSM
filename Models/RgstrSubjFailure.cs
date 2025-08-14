using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class RgstrSubjFailure
    {

        public int ROW_KEY { get; set; }

        public int SPKEY { get; set; }

        public string STATUS_INFO { get; set; }

        public string ORIGSTATUS_INFO { get; set; }

        public string SITEID { get; set; }

        public string ORIGSITEID { get; set; }

        public string SUBJID { get; set; }

        public string ORIGSUBJID { get; set; }

        public string BRTHDTC { get; set; }

        public string ORIGBRTHDTC { get; set; }

        public string SEX { get; set; }

        public string ORIGSEX { get; set; }

        public string ICDTC { get; set; }

        public string ORIGICDTC { get; set; }

        public string SCRNDTC { get; set; }

        public string AgeGroup { get; set; }

        public string ELIGRAND { get; set; }

        public string StratumID { get; set; }

        public string StratumCode { get; set; }

        public string SUBJAGE { get; set; }

        public string SELINV { get; set; }

        public string RANDNUM { get; set; }

        public string ARMCD { get; set; }

        public string ARM { get; set; }

        public string RANDBY { get; set; }

        public DateTime? DATE_RAND { get; set; }

        public string SFDATE { get; set; }

        public DateTime? SFDTC { get; set; }

        public string PM_COMM { get; set; }

        public string AMAREX_COMM { get; set; }

        public string ADDUSER { get; set; }

        public DateTime? ADDDATE { get; set; }

        public string CHANGEUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }

        public string UserName { get; set; }

        public string Password { get; set; }
    }
}
