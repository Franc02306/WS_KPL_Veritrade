using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Text;
using System.Text.RegularExpressions;

namespace Veritrade_Logistica_Service
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly IConfiguration _configuration;
        private readonly GraphServiceClient _graphClient;

        public Worker(ILogger<Worker> logger, IConfiguration configuration, GraphServiceClient graphClient)
        {
            _logger = logger;
            _configuration = configuration;
            _graphClient = graphClient;
        }

        /// <summary>
        /// Ejecutar el servicio en segundo plano
        /// </summary>
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    await ProcessInputAsync(stoppingToken);
                }
                catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error consultando Veritrade Logistica");
                }

                try
                {
                    await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken);
                }
                catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
                {
                    break;
                }
            }
        }

        /// <summary>
        /// Consultar la carpeta Input y recorrer sus subcarpetas
        /// </summary>
        private async Task ProcessInputAsync(CancellationToken cancellationToken)
        {
            var inputUrl = _configuration["Veritrade:OneDrive:InputUrl"] ?? throw new Exception("Falta Veritrade:OneDrive:InputUrl");
            var shareId = EncodeSharingUrl(inputUrl);

            _logger.LogInformation("Consultando Logistica/Input...");

            var folder = await _graphClient.Shares[shareId].DriveItem.GetAsync(cancellationToken: cancellationToken)
                ?? throw new Exception("No se pudo acceder a Logistica/Input.");

            var driveId = folder.ParentReference?.DriveId ?? throw new Exception("No se pudo obtener el DriveId.");
            var folderId = folder.Id ?? throw new Exception("No se pudo obtener el ItemId de Input.");

            var total = await ProcessItemsAsync(driveId, folderId, cancellationToken);

            _logger.LogInformation("Lectura finalizada. Archivos validos encontrados: {Total}", total);
        }

        /// <summary>
        /// Recorrer archivos y subcarpetas de forma recursiva
        /// </summary>
        private async Task<int> ProcessItemsAsync(string driveId, string folderId, CancellationToken cancellationToken, string relativePath = "")
        {
            int total = 0;
            var page = await _graphClient.Drives[driveId].Items[folderId].Children.GetAsync(cancellationToken: cancellationToken);

            while (page is not null)
            {
                foreach (var item in page.Value ?? new List<DriveItem>())
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (item.Folder is not null)
                    {
                        if (string.IsNullOrWhiteSpace(item.Id))
                            throw new Exception($"Carpeta sin ID: {item.Name}");

                        var childPath = string.IsNullOrWhiteSpace(relativePath) ? item.Name! : $"{relativePath}/{item.Name}";
                        total += await ProcessItemsAsync(driveId, item.Id, cancellationToken, childPath);

                        continue;
                    }

                    if (item.File is null || string.IsNullOrWhiteSpace(item.Name) || !item.Name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (ValidateFile(item.Name, relativePath))
                    {
                        _logger.LogInformation("Archivo encontrado: {Path}/{FileName}", relativePath, item.Name);
                        total++;
                    }
                }

                if (string.IsNullOrWhiteSpace(page.OdataNextLink))
                    break;

                page = await _graphClient.Drives[driveId].Items[folderId].Children.WithUrl(page.OdataNextLink).GetAsync(cancellationToken: cancellationToken);
            }

            return total;
        }

        /// <summary>
        /// Validar periodo, ruta y reporte configurado
        /// </summary>
        private bool ValidateFile(string fileName, string relativePath)
        {
            var match = Regex.Match(fileName, @"^(?<year>\d{4})-(?<month>0[1-9]|1[0-2])\.xlsx$", RegexOptions.IgnoreCase);

            if (!match.Success)
            {
                _logger.LogWarning("Nombre invalido: {FileName}", fileName);
                return false;
            }

            var parts = relativePath.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length != 4 || !int.TryParse(parts[0], out var folderYear) || !parts[1].Equals("Import", StringComparison.OrdinalIgnoreCase) || !Regex.IsMatch(parts[3], @"^\d{10}$"))
            {
                _logger.LogWarning("Ruta invalida: {Path}", relativePath);
                return false;
            }

            var year = int.Parse(match.Groups["year"].Value);

            if (year != folderYear)
            {
                _logger.LogWarning("El año del archivo {FileName} no coincide con su carpeta.", fileName);
                return false;
            }

            var product = parts[2];
            var tariffCode = parts[3];

            var configured = _configuration.GetSection("Veritrade:Reports").GetChildren().Any(report =>
                string.Equals(report["Operation"], "I", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(report["Product"], product, StringComparison.OrdinalIgnoreCase) &&
                report.GetSection("TariffCodes").GetChildren().Any(code => code.Value == tariffCode));

            if (!configured)
            {
                _logger.LogWarning("Reporte no configurado: {Path}", relativePath);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Codificar la URL compartida para Microsoft Graph
        /// </summary>
        private static string EncodeSharingUrl(string url)
        {
            var bytes = Encoding.UTF8.GetBytes(url);
            return "u!" + Convert.ToBase64String(bytes).TrimEnd('=').Replace('/', '_').Replace('+', '-');
        }
    }
}
