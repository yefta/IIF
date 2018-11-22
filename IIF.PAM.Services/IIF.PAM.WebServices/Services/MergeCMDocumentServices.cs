using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

using IIF.PAM.MergeDocumentServices;

namespace IIF.PAM.WebServices.Services
{
    public class MergeCMDocumentServices : BaseServices
    {
        public void DoMergeCMDocument(long id, string mergeByFQN, string mergeBy)
        {
            MergeDocument svcMerge = new MergeDocument();
            string conStringIIF = this.AppConfig.IIFConnectionString;
            svcMerge.MergeCMDocument(id, conStringIIF, this.AppConfig.CMMergeDocumentTemplateLocation, this.AppConfig.CMMergeDocumentTemporaryLocation, mergeByFQN, mergeBy);
        }
    }
}