
using SourceCode.SmartObjects.Client;

using IIF.PAM.Utilities;
using IIF.PAM.WebServices.Models;

namespace IIF.PAM.WebServices.Services
{
    public class UserADServices : BaseServices
    {
        public UserAD GetByFQN(string userFQN)
        {
            UserAD result = null;

            if (userFQN.Length > 3)
            {
                K2Services svcK2 = new K2Services();
                svcK2.AppConfig = this.AppConfig;
                SmartObjectClientServer soServer = svcK2.NewSmartObjectClientServer();

                using (soServer.Connection)
                {
                    string soName = this.AppConfig.SmartObjectName_ADUser;
                    string methodName = "GetUserDetails";
                    //load the SmartObject from the server.
                    SmartObject soAD_User = soServer.GetSmartObject(soName);

                    soAD_User.MethodToExecute = methodName;
                    //this particular method has an input parameter for the UserName
                    // Assign Input properties
                    soAD_User.Methods[methodName].InputProperties["UserName"].Value = userFQN.Right(userFQN.Length - 3);

                    soAD_User = soServer.ExecuteScalar(soAD_User);

                    result = new UserAD();
                    result.Name = soAD_User.Properties["Name"].Value;
                    result.DisplayName = soAD_User.Properties["DisplayName"].Value;
                    result.Email = soAD_User.Properties["Email"].Value;
                }
            }
            return result;
        }
    }
}