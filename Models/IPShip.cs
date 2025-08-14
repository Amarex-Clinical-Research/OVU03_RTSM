using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class IPShip
    {
        public int ILSKEY { get; set; }
        public string Courier { get; set; }

        public string TrackNo { get; set; }

        public string ExpiryDate { get; set; }

        public string LotNo { get; set; }

        public string IPLabelShip { get; set; }

        public string PMShipToDepot { get; set; }

        public string PMShipToDepotUID { get; set; }

        public DateTime? PMShipToDepotDTC { get; set; }

        public string IPLblShipReqDTC { get; set; }

        public string IPLblShipDTC { get; set; }

        public string IPLblShipUID { get; set; }

        public string IPSTAT { get; set; }

        public DateTime? ADDDATE { get; set; }

        public string ADDUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }

        public string CHANGEUSER { get; set; }
        public List<IPShipReq> IPReq { get; set; }

        public string RangeStrEnd { get; set; }
        public string DepotRangeStrEnd { get; set; }
        public string RangeStr { get; set; }

        public string RangeEnd { get; set; }
        public string IPDepotDTC { get; set; }
        public List<ShipDetails> IPDetails { get; set; }

    }
}
