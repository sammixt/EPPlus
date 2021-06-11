using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusJob
{
    public class Logger
    {
        private int LogFileSize = int.Parse("LOG_FILE_SIZE".GetKeyValue());

        public void Error(Exception ex)
        {
            if (System.IO.File.Exists("LOG_PATH".GetKeyValue() + "error" + "_log.txt"))
            {
                System.IO.FileInfo t = new System.IO.FileInfo("LOG_PATH".GetKeyValue() + "error" + "_log.txt");
                if (t.Length > LogFileSize * 1024 * 1024)
                {
                    t.MoveTo("LOG_PATH".GetKeyValue() + "error" + "_log_" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".txt");
                }
            }
            var logDetails = $"An error occurred Exception Message : {ex.Message } with stack trace : {ex.StackTrace} and Inner Message : {ex.InnerException}";
            System.IO.File.AppendAllText("LOG_PATH".GetKeyValue() + "error" + "_log.txt", DateTime.Now.ToString() + " " + logDetails + Environment.NewLine);
        }

        public void Info(string info)
        {
            if (System.IO.File.Exists("LOG_PATH".GetKeyValue() + "info" + "_log.txt"))
            {
                System.IO.FileInfo t = new System.IO.FileInfo("LOG_PATH".GetKeyValue() + "error" + "_log.txt");
                if (t.Length > LogFileSize * 1024 * 1024)
                {
                    t.MoveTo("LOG_PATH".GetKeyValue() + "info" + "_log_" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".txt");
                }
            }
            System.IO.File.AppendAllText("LOG_PATH".GetKeyValue() + "info" + "_log.txt", DateTime.Now.ToString() + " " + info + Environment.NewLine);
            Console.WriteLine($"{DateTime.Now.ToString()}:::{info}");
        }
    }
}
