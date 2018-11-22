using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class CM_WorkingPaperAttachment : BaseCMAttachment, IAttachmentType3
    {
        public string Description { get; set; }
        public int MWorkflowStatusIdWhenAdded { get; set; }
        public string SNWhenAdded { get; set; }
    }
}