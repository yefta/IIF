using System;
using System.Net;

using IIF.PAM.Utilities;

namespace IIF.PAM.SchedulerEmail.Tasks
{
    public class SendEmailOutboxTask: BaseTask
    {
        public void SendEmailOutbox()
        {
            string url = this.AppConfig.WebServiceUrl.AppendUrlPath("IIFStaticServices.asmx/SendEmailOutbox");
            HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(url);
            webrequest.Method = "GET";

            HttpWebResponse webResponse = (HttpWebResponse)webrequest.GetResponse();
            if (webResponse.StatusCode == HttpStatusCode.NotFound)
            {
                throw new Exception(url + " not found.");
            }
            else if (webResponse.StatusCode == HttpStatusCode.InternalServerError)
            {
                throw new Exception("Server encounter an error when trying to process this request.");
            }
        }
    }
}
