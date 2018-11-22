using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class CM_OtherBanksFacilitiesOrSummaryOfPefindoReport_Services : BaseAttachmentServices
    {
        public List<CM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment> ListAttachment(long cmId)
        {
            return this.ListCMAttachmentType1<CM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment>(cmId, "[dbo].[CM_OtherBanksFacilitiesOrSummaryOfPefindoReport]", "Get_CM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[CM_OtherBanksFacilitiesOrSummaryOfPefindoReport]");
        }
    }
}