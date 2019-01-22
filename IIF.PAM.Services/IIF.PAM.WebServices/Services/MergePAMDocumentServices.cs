using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

using IIF.PAM.MergeDocumentServices;

namespace IIF.PAM.WebServices.Services
{
    public class MergePAMDocumentServices : BaseServices
    {
        public void DoMergePAMDocument(long id, string mergeByFQN, string mergeBy)
        {
            MergeDocument svcMerge = new MergeDocument();
            string conStringIIF = this.AppConfig.IIFConnectionString;
            svcMerge.MergePAMDocument(id, conStringIIF, this.AppConfig.PAMMergeDocumentTemplateLocation, this.AppConfig.PAMMergeDocumentTemporaryLocation, mergeByFQN, mergeBy);

            PAM_Services svcPAM = new PAM_Services();
            svcPAM.AppConfig = this.AppConfig;
            svcPAM.DownloadPAMToSharedFolder(id);            
        }
    }
}