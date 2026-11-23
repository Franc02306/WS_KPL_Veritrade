namespace Veritrade_Service.Helpers
{
    public class SqlQueryLoader
    {
        public static async Task<string> LoadAsync(string fileName)
        {
            var path = Path.Combine(AppContext.BaseDirectory, "Queries", fileName);

            if (!File.Exists(path))
                throw new FileNotFoundException($"No se encontró el query: {path}");

            return await File.ReadAllTextAsync(path);
        }
    }
}
