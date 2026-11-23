namespace Veritrade_Logistica_Service.Helpers
{
    public static class SqlQueryLoader
    {
        /// <summary>
        /// Cargar una consulta SQL desde la carpeta Queries
        /// </summary>
        public static async Task<string> LoadAsync(string fileName, CancellationToken cancellationToken = default)
        {
            var path = Path.Combine(AppContext.BaseDirectory, "Queries", fileName);

            if (!File.Exists(path))
                throw new FileNotFoundException($"No se encontró la consulta SQL: {fileName}", path);

            return await File.ReadAllTextAsync(path, cancellationToken);
        }
    }
}
