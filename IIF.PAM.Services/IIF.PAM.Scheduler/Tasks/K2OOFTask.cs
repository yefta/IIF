using System;
using System.Net;

using IIF.PAM.Utilities;

namespace IIF.PAM.Scheduler.Tasks
{
    public class K2OOFTask: BaseTask
    {
        public void StartK2OOF()
        {
            string url = this.AppConfig.WebServiceUrl.AppendUrlPath("IIFStaticServices.asmx/StartK2OOF");
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

        public void EndK2OOF()
        {
            string url = this.AppConfig.WebServiceUrl.AppendUrlPath("IIFStaticServices.asmx/EndK2OOF");
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
