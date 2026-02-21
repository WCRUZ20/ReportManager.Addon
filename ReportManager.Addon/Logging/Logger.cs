using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportManager.Addon.Logging
{
    public sealed class Logger
    {
        private readonly string _logDir;

        public Logger(string logDir)
        {
            _logDir = logDir ?? throw new ArgumentNullException(nameof(logDir));
            Directory.CreateDirectory(_logDir);
        }

        public void Info(string message) => Write("INFO", message);
        public void Warn(string message) => Write("WARN", message);
        public void Error(string message, Exception ex = null)
            => Write("ERROR", ex == null ? message : (message + " | " + ex));

        private void Write(string level, string message)
        {
            var file = Path.Combine(_logDir, "addon_" + DateTime.Now.ToString("yyyyMMdd") + ".log");
            var line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " [" + level + "] " + message;

            // Escritura simple (suficiente para empezar)
            File.AppendAllText(file, line + Environment.NewLine, Encoding.UTF8);
        }
    }
}
