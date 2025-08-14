using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class IPLabelList
    {
        public List<KitUpload> kitupload { get; set; }
        public List<IPShip> IPShip { get; set; }
        public List<IPShipReq> IPReq { get; set; }
    }
}
