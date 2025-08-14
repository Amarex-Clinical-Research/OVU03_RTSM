using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class SPModel
    {
        public List<SiteID> Site { get; set; }
        public string shipNumber { get; set; }

        public List<ShipIP> SiteKits { get; set; }
        public List<IPStatus> DepotKits { get; set; }
    }
}
