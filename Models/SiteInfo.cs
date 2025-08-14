using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class SiteInfo
    {
        public int STSKEY { get; set; }

        public int SPKEY { get; set; }

        public int key { get; set; }
        public string SITEID { get; set; }

        public string INVNAME { get; set; }

        public string SITENAME { get; set; }

        public string ADDR1 { get; set; }

        public string ADDR2 { get; set; }

        public string CITY { get; set; }

        public string STATE { get; set; }
        public string USSTATE { get; set; }

        public string ZIPCODE { get; set; }

        public string COUNTRY { get; set; }

        public string PHONE { get; set; }

        public string FAX { get; set; }

        public string EMAIL { get; set; }

        public string SpecialInstructions { get; set; }

        public string SHIPTOEMAIL { get; set; }

        public string AMAREX_COMM { get; set; }

        public string MISC1 { get; set; }

        public string ADDUSER { get; set; }

        public DateTime? ADDDATE { get; set; }

        public string CHANGEUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }

        public string description { get; set; }
        public string Title { get; set; }

        public string Location { get; set; }
    }
}
