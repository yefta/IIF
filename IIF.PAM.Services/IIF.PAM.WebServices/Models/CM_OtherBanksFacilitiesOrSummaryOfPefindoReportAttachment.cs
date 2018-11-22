using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class CM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment : BaseCMAttachment, IAttachmentType1
    {
        public int OrderNumber { get; set; }
    }
}