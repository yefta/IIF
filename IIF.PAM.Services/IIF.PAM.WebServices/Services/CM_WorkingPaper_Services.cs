using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_WorkingPaper_Services : BaseAttachmentServices
    {
        public List<CM_WorkingPaperAttachment> ListAttachment(long cmId, int? mWorkflowStatusIdWhenAdded, int? roleIdWhenAdded, string snWhenAdded_NOT)
        {
            return this.ListCMAttachmentType3<CM_WorkingPaperAttachment>(cmId, mWorkflowStatusIdWhenAdded, roleIdWhenAdded, snWhenAdded_NOT, "[dbo].[CM_WorkingPaper]", "Get_CM_WorkingPaperAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_WorkingPaper]");
        }
    }
}