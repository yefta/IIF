
namespace IIF.PAM.WebServices.Models
{
    public class Vw_PendingReminder
    {
        public int SourceType { get; set; }
        public long SourceId { get; set; }
        public int MDocTypeId { get; set; }
        public long DocumentId { get; set; }
        public string UserFQN { get; set; }
        public string Reminder_From { get; set; }
        public string Reminder_Subject { get; set; }
        public string Reminder_Body { get; set; }
        public string Reminder_IDEmailTemplate { get; set; }
        public string Reminder_IDReference { get; set; }
    }
}