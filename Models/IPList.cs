using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class IPList
    {
        public List<IPStatus> IPStatus { get; set; }
        public List<ShipIP> ShipIP { get; set; }

        public List<DamageKit> DamageKit { get; set; }
    }
}
