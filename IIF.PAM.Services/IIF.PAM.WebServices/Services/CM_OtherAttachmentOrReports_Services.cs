using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_OtherAttachmentOrReports_Services : BaseAttachmentServices
    {
        public List<CM_OtherAttachmentOrReportsAttachment> ListAttachment(long cmId)
        {
            return this.ListCMAttachmentType1<CM_OtherAttachmentOrReportsAttachment>(cmId, "[dbo].[CM_OtherAttachmentOrReports]", "Get_CM_OtherAttachmentOrReportsAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_OtherAttachmentOrReports]");
        }
    }
}