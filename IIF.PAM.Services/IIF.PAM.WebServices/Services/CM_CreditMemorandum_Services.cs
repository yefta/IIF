using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_CreditMemorandum_Services: BaseAttachmentServices
    {
        public List<CM_CreditMemorandumAttachment> ListAttachment(long cmId)
        {
            return this.ListCMAttachmentType1<CM_CreditMemorandumAttachment>(cmId, "[dbo].[CM_CreditMemorandum]", "Get_CM_CreditMemorandumAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_CreditMemorandum]");
        }
    }
}