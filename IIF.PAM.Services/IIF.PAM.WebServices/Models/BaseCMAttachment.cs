using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class BaseCMAttachment : BaseAttachment
    {
        public long CMId { get; set; }
    }
}