using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_DocumentRelatedToBoCDecision_Services : BaseAttachmentServices
    {
        public List<PAM_DocumentRelatedToBoCDecisionAttachment> ListAttachment(long pamId, string snWhenAdded_NOT)
        {
            return this.ListPAMAttachmentType3<PAM_DocumentRelatedToBoCDecisionAttachment>(pamId, null, null, snWhenAdded_NOT, "[dbo].[PAM_DocumentRelatedToBoCDecision]", "Get_PAM_DocumentRelatedToBoCDecisionAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_DocumentRelatedToBoCDecision]");
        }
    }
}