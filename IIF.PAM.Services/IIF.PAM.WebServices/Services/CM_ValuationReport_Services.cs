using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_ValuationReport_Services : BaseAttachmentServices
    {
        public List<CM_ValuationReportAttachment> ListAttachment(long cmId)
        {
            return this.ListCMAttachmentType1<CM_ValuationReportAttachment>(cmId, "[dbo].[CM_ValuationReport]", "Get_CM_ValuationReportAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_ValuationReport]");
        }
    }
}