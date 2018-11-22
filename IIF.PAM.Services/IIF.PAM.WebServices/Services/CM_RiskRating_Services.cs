using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_RiskRating_Services: BaseAttachmentServices
    {
        public List<CM_RiskRatingAttachment> ListAttachment(long cmId)
        {
            return this.ListCMAttachmentType1<CM_RiskRatingAttachment>(cmId, "[dbo].[CM_RiskRating]", "Get_CM_RiskRatingAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_RiskRating]");
        }
    }
}