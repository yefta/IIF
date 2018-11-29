using System;
using System.Web.Services;

using IIF.PAM.WebServices.Services;

namespace IIF.PAM.WebServices
{
    /// <summary>
    /// Summary description for IIFStaticServices
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class IIFStaticServices : BaseWebService
    {
        [WebMethod]
        public string GetWebServiceUrl()
        {
            this.Logger.Info("WebServiceUrl");
            return this.AppConfig.WebServiceUrl;
            //Testing GIT
            //Testing GIT 2
        }

        [WebMethod]
        public string GetSMTPFromEmail()
        {
            return this.AppConfig.SMTPFromEmail;
        }

        [WebMethod]
        public void StartK2OOF()
        {
            try
            {
                TaskDelegationServices svcTaskDelegation = new TaskDelegationServices();
                svcTaskDelegation.AppConfig = this.AppConfig;
                svcTaskDelegation.StartK2OOF();
            }
            catch (Exception ex)
            {
                this.Logger.Error(ex);
                throw;
            }
        }

        [WebMethod]
        public void EndK2OOF()
        {
            try
            {
                TaskDelegationServices svcTaskDelegation = new TaskDelegationServices();
                svcTaskDelegation.AppConfig = this.AppConfig;
                svcTaskDelegation.EndK2OOF();
            }
            catch (Exception ex)
            {
                this.Logger.Error(ex);
                throw;
            }
        }

        [WebMethod]
        public void SendEmailOutbox()
        {
            try
            {
                OutboxEmailServices svcOutboxEmail = new OutboxEmailServices();
                svcOutboxEmail.AppConfig = this.AppConfig;
                svcOutboxEmail.SendEmailOutbox();
            }
            catch (Exception ex)
            {
                this.Logger.Error(ex);
                throw;
            }
        }

        [WebMethod]
        public void ReminderInsertOutboxEmail()
        {
            try
            {
                ReminderServices svcReminder = new ReminderServices();
                svcReminder.AppConfig = this.AppConfig;
                svcReminder.ReminderInsertOutboxEmail();
            }
            catch (Exception ex)
            {
                this.Logger.Error(ex);
                throw;
            }
        }
    }
}
