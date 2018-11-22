using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Linq;

using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class PAM_SupplementalProcurementAndInsurance_Services : BaseAttachmentServices
    {
        public List<PAM_SupplementalProcurementAndInsuranceAttachment> ListAttachment(long pamId)
        {
            return this.ListPAMAttachmentType1<PAM_SupplementalProcurementAndInsuranceAttachment>(pamId, "[dbo].[PAM_SupplementalProcurementAndInsurance]", "Get_PAM_SupplementalProcurementAndInsuranceAttachment_Content");
        }

        public XDocument GetAttachmentContent(long id)
        {
            return this.GetAttachmentContentAllType(id, "[dbo].[PAM_SupplementalProcurementAndInsurance]");
        }
    }
}