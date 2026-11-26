using System;
using System.IO;

namespace Kanna.AddIn
{
    internal static class Log
    {
        private static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Kanna", "Kanna.log");

        public static void Info(string message) => Write("INFO ", message);

        public static void Error(Exception ex) => Write("ERROR", ex.ToString());

        private static void Write(string level, string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                var fi = new FileInfo(FilePath);
                if (fi.Exists && fi.Length > 512 * 1024) fi.Delete();
                File.AppendAllText(FilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {level} {message}{Environment.NewLine}");
            }
            catch
            {
                // ログ失敗は無視
            }
        }
    }
}
