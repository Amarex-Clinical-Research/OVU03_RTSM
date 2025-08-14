using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class DamageKit
    {
        public int KITKEY { get; set; }
        public int SeqNumber { get; set; }
        public string KitNumber { get; set; }
        public string TreatmentGroup { get; set; }
        public string SITEID { get; set; }
        public short KITSET { get; set; }
        public string KITTYPE { get; set; }
        public string SENDBY { get; set; }
        public DateTime? SENDDATE { get; set; }
        public string SENDDATEstr { get; set; }
        public string RECVDDATEstr { get; set; }
        public string RECVDBY { get; set; }
        public DateTime? RECVDDATE { get; set; }
        public string KITSTAT { get; set; }
        public string ASSIGNED { get; set; }
        public DateTime? ASSIGNMENT_DATE { get; set; }
        public string ASSIGNMENT_DATEstr { get; set; }
        public string ASSIGNED_BY { get; set; }
        public string SUBJID { get; set; }
        public string USUBJID { get; set; }
        public string KITCOMM { get; set; }
        public int ROW_KEY { get; set; }
        public string VISIT { get; set; }
        public string VISITDTC { get; set; }
        public string KITINV { get; set; }
        public string KitRepled { get; set; }
        public int TELKEY { get; set; }
        public string Courier { get; set; }
        public string TrackNo { get; set; }
        public string ExpiryDate { get; set; }
        public string LotNo { get; set; }
        public int ILSKEY { get; set; }
        public string IPLblShipDTC { get; set; }
        public DateTime? IPLblShipSysDTC { get; set; }
        public string IPLblShipUID { get; set; }
        public string IPLblShipExpiryDate { get; set; }
        public string IPLblShipLotNo { get; set; }

        public string ATDEPOT { get; set; }

        public DateTime? ATDEPOTDTC { get; set; }
    }
}
