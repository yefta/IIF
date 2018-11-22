using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class CM_KYCChecklistsAttachment: BaseCMAttachment, IAttachmentType1
    {
        public int OrderNumber { get; set; }
    }
}