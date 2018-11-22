using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IIF.PAM.WebServices.Models
{
    public class OutboxEmail
    {
        public Guid Id { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string Cc { get; set; }
        public string Bcc { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Status { get; set; }
        public string IDEmailTemplate { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string IDReference { get; set; }
        public DateTime? SendDate { get; set; }
    }
}