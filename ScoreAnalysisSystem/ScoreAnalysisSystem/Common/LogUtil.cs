using log4net;

namespace ScoreAnalysisSystem.Common
{
    public class LogUtil
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(LogUtil));

        public static void Info(ILog customLogger, string message)
        {
            customLogger.Info(message);
        }

        public static void Info(string message)
        {
            Logger.Info(message);
        }

        public static void Debug(ILog customLogger, string message)
        {
            customLogger.Debug(message);
        }

        public static void Debug(string message)
        {
            Logger.Debug(message);
        }

        public static void Error(ILog customLogger, string message)
        {
            customLogger.Error(message);
        }

        public static void Error(string message)
        {
            Logger.Error(message);
        }
    }
}