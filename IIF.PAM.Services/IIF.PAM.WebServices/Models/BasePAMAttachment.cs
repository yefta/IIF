using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class BasePAMAttachment : BaseAttachment
    {
        public long PAMId { get; set; }
    }
}