using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class ScrnList
    {
        public List<SubjStat> subjList { get; set; }
        public List<Scrnstat> statList { get; set; }
        public List<ScreenFail> FailList { get; set; }
    }
}
