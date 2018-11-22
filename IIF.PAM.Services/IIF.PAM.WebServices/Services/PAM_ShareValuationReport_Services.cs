using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_ShareValuationReport_Services : BaseAttachmentServices
    {
        public List<PAM_ShareValuationReportAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType1<PAM_ShareValuationReportAttachment>(pamId, "[dbo].[PAM_ShareValuationReport]", "Get_PAM_ShareValuationReportAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_ShareValuationReport]");
        }
    }
}