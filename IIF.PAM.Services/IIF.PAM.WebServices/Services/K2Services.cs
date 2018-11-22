
using SourceCode.SmartObjects.Client;
using SourceCode.Hosting.Client.BaseAPI;

namespace IIF.PAM.WebServices.Services
{
    public class K2Services : BaseServices
    {
        public SmartObjectClientServer NewSmartObjectClientServer()
        {
            SCConnectionStringBuilder hostServerConnectionString = new SCConnectionStringBuilder();
            hostServerConnectionString.Host = this.AppConfig.K2Server;
            hostServerConnectionString.Port = 5555;
            hostServerConnectionString.IsPrimaryLogin = true;
            hostServerConnectionString.Integrated = true;
            SmartObjectClientServer result = new SmartObjectClientServer();
            result.CreateConnection();
            //open the connection to the K2 server
            result.Connection.Open(hostServerConnectionString.ToString());
            //return the SOClientServer object
            return result;
        }
    }
}