
using ClosedXML.Excel;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Text;
using System.Text.RegularExpressions;
using Veritrade_Service.Models;
using Veritrade_Service.Services;

namespace Veritrade_Service
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly IConfiguration _configuration;
        private readonly GraphServiceClient _graphClient;
        private readonly HanaService _hanaService;
        private readonly OneDriveFileService _oneDriveFileService;

        public Worker(ILogger<Worker> logger, IConfiguration configuration, GraphServiceClient graphClient, HanaService hanaService, OneDriveFileService oneDriveFileService)
        {
            _logger = logger;
            _configuration = configuration;
            _graphClient = graphClient;
            _hanaService = hanaService;
            _oneDriveFileService = oneDriveFileService;
        }

        /// <summary>
        /// Ejecutar el servicio en segundo plano
        /// </summary>
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            await _hanaService.TestConnectionAsync(stoppingToken);

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    await ProcessInputAsync(stoppingToken);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error procesando Veritrade");
                }

                await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken);
            }
        }

        /// <summary>
        /// Consultar la carpeta Input y recorrer sus subcarpetas
        /// </summary>
        private async Task ProcessInputAsync(CancellationToken cancellationToken)
        {
            var inputUrl = _configuration["Veritrade:OneDrive:InputUrl"] ?? throw new Exception("Falta Veritrade:OneDrive:InputUrl");
            var shareId = EncodeSharingUrl(inputUrl);

            _logger.LogInformation("Consultando carpeta Input...");

            var folder = await _graphClient.Shares[shareId].DriveItem.GetAsync(config =>
            {
                config.QueryParameters.Expand = new[] { "children" };
            }, cancellationToken);

            if (folder?.Children is null || folder.Children.Count == 0)
            {
                _logger.LogInformation("No hay archivos pendientes.");
                return;
            }

            var driveId = folder.ParentReference?.DriveId ?? throw new Exception("No se pudo obtener el DriveId.");

            await ProcessItemsAsync(driveId, folder.Children, cancellationToken);
        }

        /// <summary>
        /// Recorrer archivos y subcarpetas de forma recursiva
        /// </summary>
        private async Task ProcessItemsAsync(string driveId, IEnumerable<DriveItem> items, CancellationToken cancellationToken, string relativePath = "")
        {
            foreach (var item in items)
            {
                if (item.Folder is not null)
                {
                    var folderId = item.Id ?? throw new Exception($"No se pudo obtener el ItemId de la carpeta {item.Name}");
                    var children = await _graphClient.Drives[driveId].Items[folderId].Children.GetAsync(cancellationToken: cancellationToken);

                    if (children?.Value is not null && children.Value.Count > 0)
                    {
                        var childPath = string.IsNullOrWhiteSpace(relativePath) ? item.Name! : $"{relativePath}/{item.Name}";
                        await ProcessItemsAsync(driveId, children.Value, cancellationToken, childPath);
                    }

                    continue;
                }

                if (item.File is null || string.IsNullOrWhiteSpace(item.Name) || !item.Name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    continue;

                var itemId = item.Id ?? throw new Exception($"No se pudo obtener el ItemId de {item.Name}");
                var fileName = Path.GetFileNameWithoutExtension(item.Name);

                if (!Regex.IsMatch(fileName, @"^\d{4}-(0[1-9]|1[0-2])$"))
                {
                    _logger.LogWarning("Archivo omitido por nombre inválido: {FileName}", item.Name);
                    continue;
                }

                try
                {
                    await ProcessExcelAsync(driveId, item, relativePath, cancellationToken);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error procesando el archivo {FileName} en {Path}", item.Name, relativePath);

                    try
                    {
                        await _oneDriveFileService.MoveToErrorAsync(driveId, itemId, item.Name, relativePath, cancellationToken);
                    }
                    catch (Exception moveException)
                    {
                        _logger.LogError(moveException, "No se pudo mover {FileName} a Error.", item.Name);
                    }

                    continue;
                }

                try
                {
                    await _oneDriveFileService.MoveToProcessedAsync(driveId, itemId, item.Name, relativePath, cancellationToken);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "HANA fue procesado correctamente, pero no se pudo mover {FileName} a Processed.", item.Name);
                }
            }
        }

        /// <summary>
        /// Descargar y procesar un archivo Excel de Veritrade
        /// </summary>
        private async Task ProcessExcelAsync(string driveId, DriveItem item, string relativePath, CancellationToken cancellationToken)
        {
            var itemId = item.Id ?? throw new Exception($"No se pudo obtener el ItemId de {item.Name}");
            var (year, month) = GetPeriod(item.Name!);
            var (folderYear, folderOperation, folderProduct, folderTariffCode) = GetFolderContext(relativePath);

            if (year != folderYear)
                throw new Exception($"El año del archivo {item.Name} no coincide con el año de la carpeta {folderYear}.");

            _logger.LogInformation("Procesando archivo: {Path}/{FileName}", relativePath, item.Name);

            var stream = await _graphClient.Drives[driveId].Items[itemId].Content.GetAsync(cancellationToken: cancellationToken);

            if (stream is null)
                throw new Exception($"No se pudo descargar {item.Name}");

            using (stream)
            using (var workbook = new XLWorkbook(stream))
            {
                var metadataWorksheet = workbook.Worksheets.FirstOrDefault(x =>
                    x.Cell("B1").GetString().Contains("IMPORT", StringComparison.OrdinalIgnoreCase) ||
                    x.Cell("B1").GetString().Contains("EXPORT", StringComparison.OrdinalIgnoreCase))
                    ?? throw new Exception($"No se pudo identificar la operación de {item.Name}");

                var operation = GetOperation(metadataWorksheet);

                if (operation != folderOperation)
                    throw new Exception($"La operación del Excel ({operation}) no coincide con la carpeta ({folderOperation}).");

                var tariffCodes = GetTariffCodes(metadataWorksheet);

                if (tariffCodes.Count != 1)
                    throw new Exception($"El archivo {item.Name} debe contener una sola partida arancelaria. Encontradas: {string.Join(", ", tariffCodes)}");

                if (!tariffCodes[0].Equals(folderTariffCode, StringComparison.OrdinalIgnoreCase))
                    throw new Exception($"La partida del Excel ({tariffCodes[0]}) no coincide con la carpeta ({folderTariffCode}).");

                var report = GetReportConfig(operation, folderProduct, folderTariffCode);

                var worksheetName = operation == "I" ? "Importadores" : "Exportadores";
                var worksheet = workbook.Worksheet(worksheetName);
                var rows = GetRows(worksheet);

                _logger.LogInformation(
                    "Archivo identificado: {Year}-{Month:00} | {Operation} | {Product} | Partida: {TariffCode} | Hoja: {Sheet}",
                    year, month, operation, report.Product, folderTariffCode, worksheetName);

                if (operation == "I")
                    await _hanaService.ProcessImportersAsync(year, month, operation, report.Product, folderTariffCode, rows, cancellationToken);
                else
                    await _hanaService.ProcessExportersAsync(year, month, operation, report.Product, folderTariffCode, rows, cancellationToken);

                _logger.LogInformation("{FileName} procesado correctamente. Partida: {TariffCode} | Registros insertados: {Rows}",
                    item.Name, folderTariffCode, rows.Count);
            }
        }

        /// <summary>
        /// Obtener año, operación, producto y partida desde la ruta de OneDrive
        /// </summary>
        private static (int Year, string Operation, string Product, string TariffCode) GetFolderContext(string relativePath)
        {
            var parts = relativePath.Replace('\\', '/')
                .Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            if (parts.Length != 4)
                throw new Exception($"Ruta inválida: {relativePath}. Formato esperado: Año/Import o Export/Producto/Partida.");

            if (!int.TryParse(parts[0], out var year) || year < 2000 || year > 9999)
                throw new Exception($"Año inválido en la ruta: {relativePath}");

            string operation;

            if (parts[1].Equals("Import", StringComparison.OrdinalIgnoreCase))
                operation = "I";
            else if (parts[1].Equals("Export", StringComparison.OrdinalIgnoreCase))
                operation = "E";
            else
                throw new Exception($"Operación inválida en la ruta: {parts[1]}");

            var product = parts[2].ToUpperInvariant();
            var tariffCode = parts[3];

            if (!Regex.IsMatch(tariffCode, @"^\d{10}$"))
                throw new Exception($"Partida arancelaria inválida en la ruta: {tariffCode}");

            return (year, operation, product, tariffCode);
        }

        /// <summary>
        /// Obtener año y mes desde el nombre YYYY-MM.xlsx
        /// </summary>
        private static (int Year, int Month) GetPeriod(string fileName)
        {
            var name = Path.GetFileNameWithoutExtension(fileName);
            var parts = name.Split('-');

            if (parts.Length != 2 || !int.TryParse(parts[0], out var year) || !int.TryParse(parts[1], out var month) || month is < 1 or > 12)
                throw new Exception($"Nombre inválido: {fileName}. Formato esperado: YYYY-MM.xlsx");

            return (year, month);
        }

        /// <summary>
        /// Obtener tipo de operación desde la celda B1
        /// </summary>
        private static string GetOperation(IXLWorksheet worksheet)
        {
            var value = worksheet.Cell("B1").GetString();

            if (value.Contains("IMPORT", StringComparison.OrdinalIgnoreCase))
                return "I";

            if (value.Contains("EXPORT", StringComparison.OrdinalIgnoreCase))
                return "E";

            throw new Exception("No se pudo identificar la operación del Excel.");
        }

        /// <summary>
        /// Obtener partidas arancelarias desde la celda B2
        /// </summary>
        private static List<string> GetTariffCodes(IXLWorksheet worksheet)
        {
            var value = worksheet.Cell("B2").GetString();

            var codes = Regex.Matches(value, @"\b\d{10}\b").Select(x => x.Value).Distinct().OrderBy(x => x).ToList();

            if (codes.Count == 0)
                throw new Exception("No se encontraron partidas arancelarias.");

            return codes;
        }

        /// <summary>
        /// Identificar la configuración del reporte por operación, producto y partida
        /// </summary>
        private VeritradeReportConfig GetReportConfig(string operation, string product, string tariffCode)
        {
            var reports = _configuration.GetSection("Veritrade:Reports").Get<List<VeritradeReportConfig>>()
                ?? throw new Exception("Falta Veritrade:Reports");

            return reports.FirstOrDefault(x =>
                x.Operation.Equals(operation, StringComparison.OrdinalIgnoreCase) &&
                x.Product.Equals(product, StringComparison.OrdinalIgnoreCase) &&
                x.TariffCodes.Any(code => code.Equals(tariffCode, StringComparison.OrdinalIgnoreCase)))
                ?? throw new Exception($"No existe configuración para operación {operation}, producto {product} y partida {tariffCode}");
        }

        /// <summary>
        /// Codificar la URL compartida para Microsoft Graph
        /// </summary>
        private static string EncodeSharingUrl(string url)
        {
            var bytes = Encoding.UTF8.GetBytes(url);
            return "u!" + Convert.ToBase64String(bytes).TrimEnd('=').Replace('/', '_').Replace('+', '-');
        }

        /// <summary>
        /// Obtener líneas de datos desde la hoja de Excel, ignorando filas vacías y totales
        /// </summary>
        private static List<VeritradeDataRow> GetRows(IXLWorksheet worksheet)
        {
            var rows = new List<VeritradeDataRow>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

            for (var rowNumber = 7; rowNumber <= lastRow; rowNumber++)
            {
                var company = worksheet.Cell(rowNumber, 1).GetString().Trim();

                if (string.IsNullOrWhiteSpace(company) || company.Equals("Total", StringComparison.OrdinalIgnoreCase))
                    continue;

                worksheet.Cell(rowNumber, 3).TryGetValue(out decimal totalCifUsd);
                worksheet.Cell(rowNumber, 5).TryGetValue(out decimal totalKg);

                rows.Add(new VeritradeDataRow
                {
                    Company = company,
                    TotalCifUsd = totalCifUsd,
                    TotalKg = totalKg
                });
            }

            return rows;
        }
    }
}
