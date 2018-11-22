using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_LegalDueDiligenceReport_Services : BaseAttachmentServices
    {
        public List<PAM_LegalDueDiligenceReportAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType2<PAM_LegalDueDiligenceReportAttachment>(pamId, "[dbo].[PAM_LegalDueDiligenceReport]", "Get_PAM_LegalDueDiligenceReportAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_LegalDueDiligenceReport]");
        }
    }
}