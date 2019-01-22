using System;
using System.Web.Services;

using IIF.PAM.WebServices.Models;
using IIF.PAM.WebServices.Services;

namespace IIF.PAM.WebServices
{
    /// <summary>
    /// Summary description for IIFPAMServices
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class IIFPAMServices : BaseWebService
    {
        [WebMethod]
        public void DoMergePAMDocument(long id, string mergeByFQN, string mergeBy)
        {
            try
            {
                MergePAMDocumentServices svcMerge = new MergePAMDocumentServices();
                svcMerge.AppConfig = this.AppConfig;
                svcMerge.DoMergePAMDocument(id, mergeByFQN, mergeBy);
            }
            catch (Exception ex)
            {
                this.Logger.Error(ex);
                throw;
            }
        }

        [WebMethod]
        public void DownloadPAMToSharedFolder(long id)
        {
            try
            {
                PAM_Services svcPAM = new PAM_Services();
                svcPAM.AppConfig = this.AppConfig;
                svcPAM.DownloadPAMToSharedFolder(id);  
            }
            catch (Exception ex)
            {
                this.Logger.Error(ex);
                throw;
            }
        }

        [WebMethod]
        public void SendGroupEmail(PAMGroupEmailParameter param)
        {
            try
            {
                OutboxEmailServices svcOutboxEmail = new OutboxEmailServices();
                svcOutboxEmail.AppConfig = this.AppConfig;
                svcOutboxEmail.InsertPAMGroupEmail(param);
            }
            catch (Exception ex)
            {
                this.Logger.Error(ex);
                throw;
            }
        }
    }
}