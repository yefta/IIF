using System.Collections.Generic;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_OtherReports_Services : BaseAttachmentServices
    {
        public List<PAM_OtherReportsAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType2<PAM_OtherReportsAttachment>(pamId, "[dbo].[PAM_OtherReports]", "Get_PAM_OtherReportsAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_OtherReports]");
        }
    }
}