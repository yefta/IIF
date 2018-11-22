using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_OtherSupportingDocument_Services : BaseAttachmentServices
    {
        public List<CM_OtherSupportingDocumentAttachment> ListAttachment(long cmId, int? mWorkflowStatusIdWhenAdded, int? roleIdWhenAdded, string snWhenAdded_NOT)
        {
            return this.ListCMAttachmentType3<CM_OtherSupportingDocumentAttachment>(cmId, mWorkflowStatusIdWhenAdded, roleIdWhenAdded, snWhenAdded_NOT, "[dbo].[CM_OtherSupportingDocument]", "Get_CM_OtherSupportingDocumentAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_OtherSupportingDocument]");
        }
    }
}