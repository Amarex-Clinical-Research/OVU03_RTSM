using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class Visits
    {
        public List<IPStatus> IPVisit { get; set; }
        public List<VisitHistory> Visit { get; set; }
        public int VISITKEY { get; set; }

        public int SPKEY { get; set; }

        public int ROW_KEY { get; set; }

        public string SITEID { get; set; }

        public string SUBJID { get; set; }

        public string VISITID { get; set; }
        public string VISITDTC { get; set; }
        public bool notdone { get; set; }

        public string VISIT { get; set; }

        public string ELIGIP { get; set; }

        public string ReaNo { get; set; }

        public string KitNumber { get; set; }

        public string ARM { get; set; }

        public string VISITCOMM { get; set; }

        public string ReasonRep { get; set; }

        public string ADDUSER { get; set; }

        public DateTime? ADDDATE { get; set; }

        public string CHANGEUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }
        public string BRTHDTC { get; set; }
        public string SEX { get; set; }
       
        public string ReaYes { get; set; }
        public string VisitDate { get; set; }
    }
}
