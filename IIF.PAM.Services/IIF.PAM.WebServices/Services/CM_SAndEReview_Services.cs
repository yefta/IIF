using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_SAndEReview_Services : BaseAttachmentServices
    {
        public List<CM_SAndEReviewAttachment> ListAttachment(long cmId)
        {
            return this.ListCMAttachmentType1<CM_SAndEReviewAttachment>(cmId, "[dbo].[CM_SAndEReview]", "Get_CM_SAndEReviewAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_SAndEReview]");
        }
    }
}