using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

using log4net;
using log4net.Config;

using IIF.PAM.SchedulerEmail.Models;
using IIF.PAM.SchedulerEmail.Properties;
using IIF.PAM.SchedulerEmail.Tasks;

namespace IIF.PAM.SchedulerEmail
{
    class Program
    {
        private static ILog _Logger = null;
        private static ILog Logger
        {
            get
            {
                if (_Logger == null)
                {
                    _Logger = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
                    XmlConfigurator.Configure();
                }
                return _Logger;
            }
        }

        static void Main(string[] args)
        {
            log4net.Config.XmlConfigurator.Configure();

            Logger.Info("Scheduler - Start");

            ApplicationConfig appConfig = new ApplicationConfig();            
            appConfig.WebServiceUrl = Settings.Default.WebServiceUrl;
            
            try
            {
                SendEmailOutboxTask tskSendEmailOutbox = new SendEmailOutboxTask();
                tskSendEmailOutbox.AppConfig = appConfig;

                Logger.Info("Send OutboxEmail - Start"); 
                tskSendEmailOutbox.SendEmailOutbox();
                Logger.Info("Send OutboxEmail - Done"); 

                Logger.Info("Scheduler - Done");
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }

            //Console.ReadLine();
        }
    }
}
