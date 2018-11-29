﻿using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_TermSheet_Services : BaseAttachmentServices
    {
        public List<PAM_TermSheetAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType1<PAM_TermSheetAttachment>(pamId, "[dbo].[PAM_TermSheet]", "Get_PAM_TermSheetAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_TermSheet]");
        }
    }
}