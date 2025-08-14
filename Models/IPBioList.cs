using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class IPBioList
    {
        public List<KitUpload> kitupload { get; set; }
        public List<RandUpload> randupload { get; set; }

        public List<RemovedKits> removed { get; set; }
        public Byte[] UplData { get; set; }


        public IFormFile File { get; set; }
        public string Comment { get; set; }

        public string Reason { get; set; }
        public string EmailAddress { get; set; }

    }
}
