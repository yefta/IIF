using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_KYCChecklists_Services : BaseAttachmentServices
    {
        public List<CM_KYCChecklistsAttachment> ListAttachment(long cmId)
        {
            return this.ListCMAttachmentType1<CM_KYCChecklistsAttachment>(cmId, "[dbo].[CM_KYCChecklists]", "Get_CM_KYCChecklistsAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_KYCChecklists]");
        }
    }
}