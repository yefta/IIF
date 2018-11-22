using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_SAndEDueDiligence_Services : BaseAttachmentServices
    {
        public List<PAM_SAndEDueDiligenceAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType2<PAM_SAndEDueDiligenceAttachment>(pamId, "[dbo].[PAM_SAndEDueDiligence]", "Get_PAM_SAndEDueDiligenceAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_SAndEDueDiligence]");
        }
    }
}