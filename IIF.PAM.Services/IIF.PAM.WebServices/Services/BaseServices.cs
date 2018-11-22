using System.Data;
using System.Data.SqlClient;
using System.Reflection;

using log4net;

namespace IIF.PAM.WebServices.Services
{
    public class BaseServices
    {
        private ILog _Logger = null;
        protected ILog Logger
        {
            get
            {
                if (_Logger == null)
                {
                    _Logger = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
                    log4net.Config.XmlConfigurator.Configure();
                }
                return _Logger;
            }
        }

        public ApplicationConfig AppConfig { get; set; }

        protected SqlParameter NewSqlParameter(string parameterName, SqlDbType dbType, object value)
        {
            SqlParameter result = new SqlParameter(parameterName, dbType);
            result.Value = value;
            return result;
        }

    }
}