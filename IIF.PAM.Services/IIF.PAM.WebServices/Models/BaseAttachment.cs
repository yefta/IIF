using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IIF.PAM.WebServices.Models
{
    public class BaseAttachment
    {
        public string FullFileName { get; set; }
        public long Id { get; set; }
        public string DownloadHyperlink { get; set; }

        public string CreatedByFQN { get; set; }
        public string CreatedBy { get; set; }
        public DateTime CreatedOn { get; set; }
        public string ModifiedByFQN { get; set; }
        public string ModifiedBy { get; set; }
        public DateTime ModifiedOn { get; set; }

    }
}