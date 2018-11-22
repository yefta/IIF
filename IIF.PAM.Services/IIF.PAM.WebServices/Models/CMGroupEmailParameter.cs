using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class CMGroupEmailParameter : BaseGroupEmailParameter
    {
        public long CMId { get; set; }
    }
}