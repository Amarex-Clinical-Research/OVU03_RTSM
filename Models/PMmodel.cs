using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class PMmodel
    {
        public List<Randstat> randreq { get; set; }
        public List<ScreenInfo> screenreq { get; set; }
        public List<ProfileInfo> stopat { get; set; }

        public List<SiteRandStatus> SiteRand { get; set; }
        public List<SiteScreenStatus> SiteScreen { get; set; }
        public List<StopScreenAt> StopScreenAt { get; set; }
    }
}
