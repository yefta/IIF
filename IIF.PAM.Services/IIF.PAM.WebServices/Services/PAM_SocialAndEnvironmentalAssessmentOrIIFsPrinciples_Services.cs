using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciples_Services : BaseAttachmentServices
    {
        public List<PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciplesAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType1<PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciplesAttachment>(pamId, "[dbo].[PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciples]", "Get_PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciplesAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_SocialAndEnvironmentalAssessmentOrIIFsPrinciples]");
        }
    }
}