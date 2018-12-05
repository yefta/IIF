using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.Web;
using System.Web.Configuration;
using System.Web.Services;
using System.Xml.Linq;

using log4net;

namespace IIF.PAM.WebServices
{
    public class BaseWebService : WebService
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

        private ApplicationConfig _appConfig;
        protected ApplicationConfig AppConfig
        {
            get
            {
                if (this._appConfig == null)
                {
                    this._appConfig = new ApplicationConfig();                    
                    this._appConfig.IIFConnectionString = ConfigurationManager.AppSettings["IIFConnectionString"];
                    this._appConfig.WebServiceUrl = ConfigurationManager.AppSettings["WebServiceUrl"];
                    this._appConfig.ReportViewerUrl = ConfigurationManager.AppSettings["ReportViewerUrl"];
                    this._appConfig.K2Server = ConfigurationManager.AppSettings["K2Server"];
                    this._appConfig.PAMMergeDocumentTemplateLocation = ConfigurationManager.AppSettings["PAMMergeDocumentTemplateLocation"];
                    this._appConfig.PAMMergeDocumentTemporaryLocation = ConfigurationManager.AppSettings["PAMMergeDocumentTemporaryLocation"];
                    this._appConfig.CMMergeDocumentTemplateLocation = ConfigurationManager.AppSettings["CMMergeDocumentTemplateLocation"];
                    this._appConfig.CMMergeDocumentTemporaryLocation = ConfigurationManager.AppSettings["CMMergeDocumentTemporaryLocation"];
                    this._appConfig.SmartObjectName_ADUser = ConfigurationManager.AppSettings["SmartObjectName_ADUser"];

                    this._appConfig.SMTPFromEmail = ConfigurationManager.AppSettings["SMTPFromEmail"];
                    this._appConfig.SMTPFromName = ConfigurationManager.AppSettings["SMTPFromName"];
                    this._appConfig.SMTPHost = ConfigurationManager.AppSettings["SMTPHost"];
                    this._appConfig.SMTPPort = int.Parse(ConfigurationManager.AppSettings["SMTPPort"]);
                    this._appConfig.SMTPEnableSSL = ConfigurationManager.AppSettings["SMTPEnableSSL"].ToUpper() == "TRUE";
                    this._appConfig.SMTPCredentialName = ConfigurationManager.AppSettings["SMTPCredentialName"];
                    this._appConfig.SMTPCredentialPassword = ConfigurationManager.AppSettings["SMTPCredentialPassword"];
                }
                return this._appConfig;
            }
        }

        protected SqlParameter NewSqlParameter(string parameterName, SqlDbType dbType, object value)
        {
            SqlParameter result = new SqlParameter(parameterName, dbType);
            result.Value = value;
            return result;
        }

        protected void ReturnXDocumentFile(XDocument xDoc)
        {
            byte[] byteContent = Convert.FromBase64String(xDoc.Root.Element("content").Value);
            string fullFileName = xDoc.Root.Element("name").Value;

            HttpResponse response = this.Context.Response;
            response.Clear();
            response.ClearHeaders();
            response.ContentType = "Application/octet-stream";
            response.AddHeader("Content-Disposition", "attachment; filename=\"" + fullFileName + "\"");
            response.OutputStream.Write(byteContent, 0, byteContent.Length);
            response.Flush();
            response.End();
        }
    }
}