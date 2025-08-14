using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class Email_Notifications
    {
        public List<RandEmailNotif> randemail { get; set; }
        public List<ShipperEmail> shipemail { get; set; }

        public List<BiometircUploadEmailNoti> Bioemail { get; set; }

        public List<IPLabRelShip> Regulatory { get; set; }

        public List<IPlabRelUp> EmergencyUnblind { get; set; }

        public int KEYID { get; set; }

        public int SPKEY { get; set; }

        public string PIDESC { get; set; }

        public string PIDet { get; set; }

        public string PIType { get; set; }

        public DateTime? ADDDATE { get; set; }

        public string ADDUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }

        public string CHANGEUSER { get; set; }
        public bool Enable { get; set; }


    }
}
