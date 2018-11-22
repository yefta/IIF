using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class CM_PeriodicReviewAttachment : BaseCMAttachment, IAttachmentType1
    {
        public int OrderNumber { get; set; }
    }
}