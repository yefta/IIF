using System;
using System.Web.Services;

using IIF.PAM.WebServices.Models;
using IIF.PAM.WebServices.Services;

namespace IIF.PAM.WebServices
{
    /// <summary>
    /// Summary description for IIFCMServices
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class IIFCMServices : BaseWebService
    {
		[WebMethod]
		public void DoMergePAMDocument(long id, string mergeByFQN, string mergeBy)
        {
            try
            {
                MergeCMDocumentServices svcMerge = new MergeCMDocumentServices();
                svcMerge.AppConfig = this.AppConfig;
                //svcMerge.DoMergeCMDocument(id, mergeByFQN, mergeBy);
            }
            catch (Exception ex)
            {
                this.Logger.Error(ex);
                throw;
            }
        }

		[WebMethod]
		public void SendGroupEmail(CMGroupEmailParameter param)
        {
            try
            {
                OutboxEmailServices svcOutboxEmail = new OutboxEmailServices();
                svcOutboxEmail.AppConfig = this.AppConfig;
                svcOutboxEmail.InsertCMGroupEmail(param);
            }
            catch (Exception ex)
            {
                this.Logger.Error(ex);
                throw;
            }
        }
    }
}
