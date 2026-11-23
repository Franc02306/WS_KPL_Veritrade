namespace Veritrade_Service.Logging
{
    public class FileLogger : ILogger
    {
        private static readonly object Lock = new();
        private readonly string _logsPath;
        private readonly string _categoryName;

        public FileLogger(string logsPath, string categoryName)
        {
            _logsPath = logsPath;
            _categoryName = categoryName;
        }

        /// <summary>
        /// Inicia un scope de registro
        /// </summary>
        public IDisposable? BeginScope<TState>(TState state) where TState : notnull
        {
            return null;
        }

        /// <summary>
        /// Activado para registrar mensajes de nivel Information y superior (Warning, Error, Critical)
        /// </summary>
        public bool IsEnabled(LogLevel logLevel)
        {
            return logLevel >= LogLevel.Information;
        }

        /// <summary>
        /// Logear un mensaje en un archivo de registro
        /// </summary>
        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
        {
            if (!IsEnabled(logLevel))
                return;

            var filePath = Path.Combine(_logsPath, $"veritrade-{DateTime.Now:yyyyMMdd}.log");
            var message = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{logLevel}] {_categoryName} - {formatter(state, exception)}";

            if (exception is not null)
                message += Environment.NewLine + exception;

            lock (Lock)
            {
                File.AppendAllText(filePath, message + Environment.NewLine);
            }
        }
    }
}
