using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class BiometircUploadEmailNoti
    {
        public int KEYID { get; set; }

        public int SPKEY { get; set; }

        public string PIDESC { get; set; }

        public string PIDet { get; set; }

        public string PIType { get; set; }

        public DateTime? ADDDATE { get; set; }

        public string ADDUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }

        public string CHANGEUSER { get; set; }
        public int UPLOAD_KEY { get; set; }
        public bool Enable { get; set; }

        public string Comment { get; set; }
        public string Reason { get; set; }
    }
}
