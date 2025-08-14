using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class RemovedKits
    {
        public int UPLOAD_KEY { get; set; }

        public int SPKEY { get; set; }

        public string UPLOAD_DESC { get; set; }

        public string FILENAME { get; set; }

        public string EMAILSENT { get; set; }

        public string TYPEUL { get; set; }

        public string ADDUSER { get; set; }

        public DateTime? ADDDATE { get; set; }

        public string CHANGEUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }

        public string IsHide { get; set; }

        public string Reason { get; set; }

        public string EmailAddress { get; set; }

        public string CodeSend { get; set; }
    }
}
