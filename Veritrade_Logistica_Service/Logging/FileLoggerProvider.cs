namespace Veritrade_Logistica_Service.Logging
{
    public class FileLoggerProvider : ILoggerProvider
    {
        private readonly string _logsPath;

        public FileLoggerProvider()
        {
            _logsPath = Path.Combine(AppContext.BaseDirectory, "Logs");
            Directory.CreateDirectory(_logsPath);
        }

        /// <summary>
        /// Crear un logger para la categoría especificada
        /// </summary>
        public ILogger CreateLogger(string categoryName)
        {
            return new FileLogger(_logsPath, categoryName);
        }

        public void Dispose()
        {
        }
    }
}
