using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class InventoryModel
    {
        public List<StopAutoSupply> Stopauto { get; set; }
        public List<Inventory> Inventory { get; set; }
        public List<SiteID> Site { get; set; }
        public List<DepotInv> Depinv { get; set; }
        public List<StopAutoSupply> StudyResupply { get; set; }
    }
}
