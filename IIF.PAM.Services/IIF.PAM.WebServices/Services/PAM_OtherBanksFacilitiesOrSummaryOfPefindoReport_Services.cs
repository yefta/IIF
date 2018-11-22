using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_OtherBanksFacilitiesOrSummaryOfPefindoReport_Services : BaseAttachmentServices
    {
        public List<PAM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType1<PAM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment>(pamId, "[dbo].[PAM_OtherBanksFacilitiesOrSummaryOfPefindoReport]", "Get_PAM_OtherBanksFacilitiesOrSummaryOfPefindoReportAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_OtherBanksFacilitiesOrSummaryOfPefindoReport]");
        }
    }
}