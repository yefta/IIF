using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class IIFWorklistItem_Filter
    {
        public string Destination { get; set; }
        public DateTime? SubmitDate_FROM { get; set; }
        public DateTime? SubmitDate_TO { get; set; }
        public string ProjectCode_LIKE { get; set; }
        public string CustomerName_LIKE { get; set; }
        public int? ProductTypeId { get; set; }
        public int? MDocTypeId { get; set; }
        public string CMNumber_LIKE { get; set; }
    }
}