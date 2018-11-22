using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IIF.PAM.WebServices.Models
{
    public class TaskDelegation
    {
        public long Id { get; set; }
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string FromFQN { get; set; }
        public string ToFQN { get; set; }
        public bool IsActive { get; set; }
        public bool IsExpired { get; set; }
        public bool IsCanceled { get; set; }
        public bool IsStartedInK2 { get; set; }
        public bool IsEndedInK2 { get; set; }
    }
}