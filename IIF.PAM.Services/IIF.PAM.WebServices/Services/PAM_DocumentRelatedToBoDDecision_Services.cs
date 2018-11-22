using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_DocumentRelatedToBoDDecision_Services : BaseAttachmentServices
    {
        public List<PAM_DocumentRelatedToBoDDecisionAttachment> ListAttachment(long pamId, string snWhenAdded_NOT)
        {
            return this.ListPAMAttachmentType3<PAM_DocumentRelatedToBoDDecisionAttachment>(pamId, null, null, snWhenAdded_NOT, "[dbo].[PAM_DocumentRelatedToBoDDecision]", "Get_PAM_DocumentRelatedToBoDDecisionAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_DocumentRelatedToBoDDecision]");
        }
    }
}