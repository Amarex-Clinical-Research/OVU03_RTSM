using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class IPShipReq
    {

        public int TELKEY { get; set; }

        public int ILSKEY { get; set; }

        public string Courier { get; set; }

        public string TrackNo { get; set; }

        public string ExpiryDate { get; set; }

        public string LotNo { get; set; }

        public string IPLabelShip { get; set; }

        public string RangeStr { get; set; }

        public string RangeEnd { get; set; }

        public DateTime? ADDDATE { get; set; }

        public string ADDUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }

        public string CHANGEUSER { get; set; }

        public string IPDepotDTC { get; set; }
    }
}
