using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Webview_IRT.Models
{
    public class IPUploadFiles
    {
        public int UPLOAD_KEY { get; set; }

        public int SPKEY { get; set; }

        public string FILENAME { get; set; }

        public string TYPEUL { get; set; }

        public byte[] DataFile { get; set; }

        public string FUSTATUS { get; set; }

        public string PM_COMM { get; set; }

        public string AMAREX_COMM { get; set; }

        public string ADDUSER { get; set; }

        public DateTime? ADDDATE { get; set; }

        public string CHANGEUSER { get; set; }

        public DateTime? CHANGEDATE { get; set; }
        public IFormFile PackageFile { get; set; }
        public IFormFile LogFile { get; set; }
        public short KITSET { get; set; }
        public string SITEID { get; set; }

    }
}
