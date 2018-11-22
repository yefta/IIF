using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_GroupStructure_Services : BaseAttachmentServices
    {
        public List<PAM_GroupStructureAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType1<PAM_GroupStructureAttachment>(pamId, "[dbo].[PAM_GroupStructure]", "Get_PAM_GroupStructureAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_GroupStructure]");
        }
    }
}