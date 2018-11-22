using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_IndustryAnalysis_Services : BaseAttachmentServices
    {
        public List<PAM_IndustryAnalysisAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType1<PAM_IndustryAnalysisAttachment>(pamId, "[dbo].[PAM_IndustryAnalysis]", "Get_PAM_IndustryAnalysisAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_IndustryAnalysis]");
        }
    }
}