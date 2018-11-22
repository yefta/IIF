using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class PAMGroupEmailParameter: BaseGroupEmailParameter
    {
        public long PAMId { get; set; }
        
    }
}