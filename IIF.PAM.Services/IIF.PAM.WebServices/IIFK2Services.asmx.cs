using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;

using IIF.PAM.WebServices.Models;
using IIF.PAM.WebServices.Services;

namespace IIF.PAM.WebServices
{
    /// <summary>
    /// Summary description for IIFK2Services
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class IIFK2Services : BaseWebService
    {
        [WebMethod]
        public List<IIFWorklistItem> ListIIFWorklistItem(IIFWorklistItem_Filter filter)
        {
            K2Services svc = new K2Services();
            svc.AppConfig = this.AppConfig;
            return svc.ListIIFWorklistItem(filter);
        }

        [WebMethod]
        public void RetryWorkflow()
        {
            K2Services svc = new K2Services();
            svc.AppConfig = this.AppConfig;
            svc.RetryWorkflow();
        }

    }
}
