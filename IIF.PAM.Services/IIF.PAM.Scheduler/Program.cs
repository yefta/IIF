using System;
using System.Reflection;

using IIF.PAM.Scheduler.Models;
using IIF.PAM.Scheduler.Properties;
using IIF.PAM.Scheduler.Tasks;

using log4net;

namespace IIF.PAM.Scheduler
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
                    log4net.Config.XmlConfigurator.Configure();
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

            ReminderTask tskReminder = new ReminderTask();
            tskReminder.AppConfig = appConfig;

            K2OOFTask tskK2OOF = new K2OOFTask();
            tskK2OOF.AppConfig = appConfig;

            try
            {
                Logger.Info("Send Reminder Email - Start");
                tskReminder.SendReminderEmail();
                Logger.Info("Send Reminder Email - Done");
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }

            try
            {
                Logger.Info("End K2 Out of Office - Start");
                tskK2OOF.EndK2OOF();
                Logger.Info("End K2 Out of Office - Done");
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }

            try
            {
                Logger.Info("Start K2 Out of Office - Start");
                tskK2OOF.StartK2OOF();
                Logger.Info("Start K2 Out of Office - Done");
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
            Logger.Info("Scheduler - Done");
        }
    }
}