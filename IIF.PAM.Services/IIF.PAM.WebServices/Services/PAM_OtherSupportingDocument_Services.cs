using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_OtherSupportingDocument_Services : BaseAttachmentServices
    {
        public List<PAM_OtherSupportingDocumentAttachment> ListAttachment(long pamId, int? mWorkflowStatusIdWhenAdded, int? roleIdWhenAdded, string snWhenAdded_NOT)
        {
            return this.ListPAMAttachmentType3<PAM_OtherSupportingDocumentAttachment>(pamId, mWorkflowStatusIdWhenAdded, roleIdWhenAdded, snWhenAdded_NOT, "[dbo].[PAM_OtherSupportingDocument]", "Get_PAM_OtherSupportingDocumentAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_OtherSupportingDocument]");
        }
    }
}