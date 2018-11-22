using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class PAM_SAndEDueDiligenceAttachment : BasePAMAttachment, IAttachmentType2
    {
        public string Description { get; set; }
    }
}