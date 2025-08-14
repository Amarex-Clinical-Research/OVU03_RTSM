using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class Audit
    {
        public string AUDTYPE { get; set; }
        public string AUDTBL { get; set; }
        public string AUDPKEY { get; set; }
        public string AUDFLD { get; set; }
        public string AUDOLDV { get; set; }
        public string AUDNEWV { get; set; }
        public string AUDUSER { get; set; }
        public DateTime AUDDTC { get; set; }
    }
}
