using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class IIFWorklistItem
    {
        public int MDocTypeId { get; set; }
        public string MDocTypeName { get; set; }
        public long DocumentId { get; set; }
        public int ProductTypeId { get; set; }
        public string ProductTypeName { get; set; }

        public string ProjectCode { get; set; }
        public string CustomerName { get; set; }
        public string CMNumber { get; set; }

        public bool IsInRevise { get; set; }

        public int WorkflowStatusId { get; set; }
        public string WorkflowStatusName { get; set; }
        public string TaskListStatus { get; set; }

        public DateTime SubmitDate { get; set; }
        public string ModifiedBy { get; set; }
        public DateTime ModifiedOn { get; set; }

        public int K2ProcessId { get; set; }
        public string SN { get; set; }
        public string K2CurrentActivityName { get; set; }
        public string SharedUserFQN { get; set; }
    }
}