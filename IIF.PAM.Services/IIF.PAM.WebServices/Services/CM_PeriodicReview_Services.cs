using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_PeriodicReview_Services: BaseAttachmentServices
    {
        public List<CM_PeriodicReviewAttachment> ListAttachment(long cmId)
        {
            return this.ListCMAttachmentType1<CM_PeriodicReviewAttachment>(cmId, "[dbo].[CM_PeriodicReview]", "Get_CM_PeriodicReviewAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_PeriodicReview]");
        }
    }
}