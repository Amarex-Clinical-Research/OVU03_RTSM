using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class Reportproblem
    {
        public int PRBLMKEY { get; set; }

        public int SPKEY { get; set; }

        public string SITEID { get; set; }

        public string SUBJID { get; set; }

        public string PROBLEM_DESC { get; set; }

        public string PM_COMM { get; set; }

        public string EMAIL_SENT { get; set; }

        public string AMAREX_COMM { get; set; }

        public string ADDUSER { get; set; }

        public string PROBLEM_STATUS { get; set; }

        public DateTime? ADDDATE { get; set; }
        public string DateReport { get; set; }

        public string CHANGEUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }

        public bool NotApp { get; set; }

        public bool SiteNA { get; set; }


    }
}
