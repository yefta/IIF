﻿using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_ProjectAnalysis_Services : BaseAttachmentServices
    {        
        public List<PAM_ProjectAnalysisAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType1<PAM_ProjectAnalysisAttachment>(pamId, "[dbo].[PAM_ProjectAnalysis]", "Get_PAM_ProjectAnalysisAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_ProjectAnalysis]");
        }
    }
}