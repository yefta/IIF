using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_RiskRating_Services : BaseAttachmentServices
    {
        public List<PAM_RiskRatingAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType1<PAM_RiskRatingAttachment>(pamId, "[dbo].[PAM_RiskRating]", "Get_PAM_RiskRatingAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_RiskRating]");
        }
    }
}