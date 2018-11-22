using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class PAM_LegalDueDiligenceReportAttachment : BasePAMAttachment, IAttachmentType2
    {
        public string Description { get; set; }
    }
}