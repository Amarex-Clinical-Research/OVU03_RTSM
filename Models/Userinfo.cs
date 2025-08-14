using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class Userinfo
    {
        public int UCKey { get; set; }

        public string UserID { get; set; }

        public int SPKey { get; set; }

        public string Center_Num { get; set; }

        public string Center_Lvl { get; set; }

        public DateTime? ADDDATE { get; set; }

        public string ADDUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }

        public string CHANGEUSER { get; set; }
    }
}
