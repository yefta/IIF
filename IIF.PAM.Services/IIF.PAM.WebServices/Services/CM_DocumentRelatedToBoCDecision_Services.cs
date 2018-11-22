using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_DocumentRelatedToBoCDecision_Services : BaseAttachmentServices
    {
        public List<CM_DocumentRelatedToBoCDecisionAttachment> ListAttachment(long CMId, string snWhenAdded_NOT)
        {
            return this.ListCMAttachmentType3<CM_DocumentRelatedToBoCDecisionAttachment>(CMId, null, null, snWhenAdded_NOT, "[dbo].[CM_DocumentRelatedToBoCDecision]", "Get_CM_DocumentRelatedToBoCDecisionAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_DocumentRelatedToBoCDecision]");
        }
    }
}