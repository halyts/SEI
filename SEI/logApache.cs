using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using log4net.Appender;
using log4net.Layout;
using log4net.Core;
[assembly: log4net.Config.XmlConfigurator(ConfigFile = "LogConfig.xml")]

namespace SEI
{
    public  class LogApache
    {
        public ILogger[] loggers;
        public IAppender[] appenders;
        public ILogger[] lgrs { get { return loggers; } }
        public IAppender[] appnds { get { return appenders; } }

        public LogApache()
        {
            appenders = LogManager.GetRepository().GetAppenders();
            loggers = LogManager.GetRepository().GetCurrentLoggers();

        }
        public static LogApache getLgApache()
        {
            return new LogApache();        
        }
    }
}
