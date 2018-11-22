using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_KYCChecklists_Services : BaseAttachmentServices
    {
        public List<PAM_KYCChecklistsAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType1<PAM_KYCChecklistsAttachment>(pamId, "[dbo].[PAM_KYCChecklists]", "Get_PAM_KYCChecklistsAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_KYCChecklists]");
        }
    }
}