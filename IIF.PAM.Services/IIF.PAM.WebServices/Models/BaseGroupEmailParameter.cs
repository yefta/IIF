using System;

namespace IIF.PAM.WebServices.Models
{
    [Serializable]
    public class BaseGroupEmailParameter
    {
        public int MRoleId { get; set; }
        public string From { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string IDEmailTemplate { get; set; }
    }
}