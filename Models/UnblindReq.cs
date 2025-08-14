using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class UnblindReq
    {
        public int UNBLDKEY { get; set; }

        public int SPKEY { get; set; }

        public int RANDKEY { get; set; }

        public bool UNBLIND_SAE { get; set; }

        public bool UNBLIND_INAPPRO { get; set; }

        public bool UNBLIND_OTH { get; set; }

        public string UNBLIND_OTH_SPEC { get; set; }

        public string UNBLIND_TERM { get; set; }

        public string UNBLIND_COMM { get; set; }

        public DateTime? UNBLIND_DATE { get; set; }

        public string UNBLIND_UID { get; set; }

        public string MMCONTACTED { get; set; }

        public DateTime? MMCONTACTDTC { get; set; }

        public string NOCONTACTREASON { get; set; }

        public string SITEID { get; set; }

        public string SUBJID { get; set; }

        public string ARM { get; set; }

        public DateTime? DATE_RAND { get; set; }

        public string AMAREX_COMM { get; set; }

        public string TYPEUNB { get; set; }

        public string ADDUSER { get; set; }

        public DateTime? ADDDATE { get; set; }

        public string CHANGEUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }

        public string BRTHDTC { get; set; }

        public string REQSTAT { get; set; }

        public string RSRESPUSER { get; set; }

        public DateTime? RSRESPDTC { get; set; }

        public string RejectSpec { get; set; }
        public string AcceptSpec { get; set; }

        public string UNBLINDDATE { get; set; }
        public string RANDDATE { get; set; }

        public string MMCONTACTDATE { get; set; }

        public string PROCESSEDDATE { get; set; }
    }
}
