using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

using log4net;
using log4net.Config;

using IIF.PAM.SchedulerRetryWorkflow.Models;
using IIF.PAM.SchedulerRetryWorkflow.Properties;
using IIF.PAM.SchedulerRetryWorkflow.Tasks;

namespace IIF.PAM.SchedulerRetryWorkflow
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
                RetryWorkflowTask tskRetryWorkflow = new RetryWorkflowTask();
                tskRetryWorkflow.AppConfig = appConfig;

                Logger.Info("Retry Workflow - Start");
                tskRetryWorkflow.RetryWorkflow();
                Logger.Info("Retry Workflow - Done");

                Logger.Info("Scheduler - Done");
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }

        }
    }
}
