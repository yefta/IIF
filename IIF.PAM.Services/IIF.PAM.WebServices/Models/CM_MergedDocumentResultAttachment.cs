using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class CM_MergedDocumentResultAttachment : BaseCMAttachment
    {
        public bool IsForHistory { get; set; }
    }
}