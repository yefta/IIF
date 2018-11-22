using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_DocumentRelatedToBoDDecision_Services : BaseAttachmentServices
    {
        public List<CM_DocumentRelatedToBoDDecisionAttachment> ListAttachment(long cmId, string snWhenAdded_NOT)
        {
            return this.ListCMAttachmentType3<CM_DocumentRelatedToBoDDecisionAttachment>(cmId, null, null, snWhenAdded_NOT, "[dbo].[CM_DocumentRelatedToBoDDecision]", "Get_CM_DocumentRelatedToBoDDecisionAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_DocumentRelatedToBoDDecision]");
        }
    }
}