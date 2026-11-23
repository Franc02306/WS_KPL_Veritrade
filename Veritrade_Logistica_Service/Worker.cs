using ClosedXML.Excel;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Text;
using System.Text.RegularExpressions;
using Veritrade_Logistica_Service.Models;
using Veritrade_Logistica_Service.Services;
using static Veritrade_Logistica_Service.Helpers.ExcelReaderHelper;

namespace Veritrade_Logistica_Service
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
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    if (_configuration.GetValue<bool>("Veritrade:EnableProcessing"))
                        await _hanaService.TestConnectionAsync(stoppingToken);

                    await ProcessInputAsync(stoppingToken);
                }
                catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error procesando Veritrade Logistica");
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
                var items = new List<DriveItem>();
                items.AddRange(page.Value ?? new List<DriveItem>());

                if (string.IsNullOrWhiteSpace(page.OdataNextLink))
                    page = null;
                else
                    page = await _graphClient.Drives[driveId].Items[folderId].Children.WithUrl(page.OdataNextLink).GetAsync(cancellationToken: cancellationToken);

                foreach (var item in items)
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

                    if (!ValidateFile(item.Name, relativePath))
                        continue;

                    var itemId = item.Id ?? throw new Exception($"No se pudo obtener el ItemId de {item.Name}");
                    total++;

                    try
                    {
                        var processed = await ProcessExcelAsync(driveId, item, relativePath, cancellationToken);

                        if (!processed)
                            continue;
                    }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Error procesando {FileName} en {Path}", item.Name, relativePath);

                        if (!_configuration.GetValue<bool>("Veritrade:EnableProcessing"))
                            continue;

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

            return total;
        }

        /// <summary>
        /// Validar periodo, ruta y reporte configurado
        /// </summary>
        private bool ValidateFile(string fileName, string relativePath)
        {
            if (!fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || fileName.StartsWith("~$"))
                return false;

            try
            {
                var (year, product, tariffCode) = GetFolderContext(relativePath);

                var configured = _configuration.GetSection("Veritrade:Reports").GetChildren().Any(report =>
                    string.Equals(report["Operation"], "I", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(report["Product"], product, StringComparison.OrdinalIgnoreCase) &&
                    report.GetSection("TariffCodes").GetChildren().Any(code => code.Value == tariffCode));

                if (!configured)
                {
                    _logger.LogWarning("Reporte no configurado: {Path}", relativePath);
                    return false;
                }

                _logger.LogInformation("Archivo identificado: {Year} | {Product} | Partida: {TariffCode} | {FileName}", year, product, tariffCode, fileName);
                return true;
            }
            catch (ArgumentException ex)
            {
                _logger.LogWarning(ex, "Ruta inválida para {FileName}: {Path}", fileName, relativePath);
                return false;
            }
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
        /// Obtiene año, producto y partida desde la ruta de OneDrive
        /// </summary>
        private static (int Year, string Product, string TariffCode) GetFolderContext(string relativePath)
        {
            var parts = relativePath.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            if (parts.Length != 4)
                throw new ArgumentException($"Ruta inválida: {relativePath}. Formato esperado: Año/Import/Producto/Partida.");

            if (!int.TryParse(parts[0], out var year) || year < 2000 || year > 9999)
                throw new ArgumentException($"Año inválido en la ruta: {relativePath}");

            if (!parts[1].Equals("Import", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException($"Operación no permitida: {parts[1]}");

            if (!Regex.IsMatch(parts[3], @"^\d{10}$"))
                throw new ArgumentException($"Partida inválida: {parts[3]}");

            return (year, parts[2].ToUpperInvariant(), parts[3]);
        }

        /// <summary>
        /// Descargar, validar y procesar el Excel anual de Veritrade
        /// </summary>
        private async Task<bool> ProcessExcelAsync(string driveId, DriveItem item, string relativePath, CancellationToken cancellationToken)
        {
            var itemId = item.Id ?? throw new Exception($"No se pudo obtener el ItemId de {item.Name}");
            var (year, product, tariffCode) = GetFolderContext(relativePath);

            _logger.LogInformation("Procesando archivo: {Path}/{FileName}", relativePath, item.Name);

            var stream = await _graphClient.Drives[driveId].Items[itemId].Content.GetAsync(cancellationToken: cancellationToken)
                ?? throw new Exception($"No se pudo descargar {item.Name}");

            using (stream)
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = workbook.Worksheets.FirstOrDefault(x => x.Cell("B1").GetString().Contains("IMPORT", StringComparison.OrdinalIgnoreCase))
                    ?? throw new Exception($"No se encontró una hoja de importaciones en {item.Name}");

                var metadataOperation = worksheet.Cell("B1").GetString();

                if (!metadataOperation.Contains("IMPORT", StringComparison.OrdinalIgnoreCase))
                    throw new Exception($"El archivo {item.Name} no corresponde a importaciones.");

                var codes = Regex.Matches(worksheet.Cell("B2").GetString(), @"(?<!\d)\d{10}(?!\d)").Select(x => x.Value).Distinct().ToList();

                if (codes.Count != 1 || codes[0] != tariffCode)
                    throw new Exception($"La partida del Excel no coincide con la carpeta: {tariffCode}");

                var periodYears = Regex.Matches(worksheet.Cell("B3").GetString(), @"\b(?:19|20)\d{2}\b")
                    .Select(x => int.Parse(x.Value)).Distinct().ToList();

                if (periodYears.Count != 1 || periodYears[0] != year)
                    throw new Exception($"El periodo del Excel no corresponde al año {year}.");

                var rows = GetRows(worksheet, year, product, tariffCode);

                if (rows.Count == 0)
                    throw new Exception($"El archivo {item.Name} no contiene registros válidos.");

                _logger.LogInformation("Excel validado: {Year} | {Product} | Partida: {TariffCode} | {Rows} registros", year, product, tariffCode, rows.Count);

                if (!_configuration.GetValue<bool>("Veritrade:EnableProcessing"))
                {
                    _logger.LogInformation("Modo prueba: no se insertaron registros en HANA ni se movió el archivo.");
                    return false;
                }

                await _hanaService.ProcessImportersAsync(year, product, tariffCode, rows, cancellationToken);
                return true;
            }
        }

        /// <summary>
        /// Obtener registros del Excel, excluyendo columnas innecesarias
        /// </summary>
        private static List<VeritradeLogisticaRow> GetRows(IXLWorksheet worksheet, int year, string product, string tariffCode)
        {
            var rows = new List<VeritradeLogisticaRow>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

            var headers = worksheet.Row(6).CellsUsed()
                .Where(x => !string.IsNullOrWhiteSpace(x.GetString()))
                .ToDictionary(x => x.GetString().Trim(), x => x.Address.ColumnNumber, StringComparer.OrdinalIgnoreCase);

            IXLCell Cell(int row, string header)
            {
                if (!headers.TryGetValue(header, out var column))
                    throw new Exception($"No se encontró la columna obligatoria: {header}");

                return worksheet.Cell(row, column);
            }

            for (var rowNumber = 7; rowNumber <= lastRow; rowNumber++)
            {
                if (worksheet.Row(rowNumber).IsEmpty())
                    continue;

                var partida = ReadText(Cell(rowNumber, "Partida Aduanera"));

                if (string.Equals(partida, "Total", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (partida != tariffCode)
                    throw new Exception($"Fila {rowNumber}: partida {partida ?? "(vacía)"} diferente de {tariffCode}.");

                var fecha = ReadDate(Cell(rowNumber, "Fecha"));

                if (fecha.HasValue && fecha.Value.Year != year)
                    throw new Exception($"Fila {rowNumber}: fecha fuera del año {year}.");

                var row = new VeritradeLogisticaRow
                {
                    Anio = year,
                    Producto = product,
                    Operacion = "I",
                    FilaExcel = rowNumber,

                    // Información aduanera
                    PartidaAduanera = partida,
                    DescripcionPartida = ReadText(Cell(rowNumber, "Descripcion de la Partida Aduanera")),
                    Aduana = ReadText(Cell(rowNumber, "Aduana")),
                    DuaDam = ReadText(Cell(rowNumber, "DUA / DAM")),
                    Fecha = fecha,
                    Eta = ReadDate(Cell(rowNumber, "ETA")),
                    CodTributario = ReadText(Cell(rowNumber, "Cod. Tributario")),
                    Importador = ReadText(Cell(rowNumber, "Importador")),
                    EmbarcadorExportador = ReadText(Cell(rowNumber, "Embarcador / Exportador")),

                    // Pesos y cantidades
                    KgBruto = ReadDecimal(Cell(rowNumber, "Kg Bruto")),
                    KgNeto = ReadDecimal(Cell(rowNumber, "Kg Neto")),
                    Qty1 = ReadDecimal(Cell(rowNumber, "Qty 1")),
                    Und1 = ReadText(Cell(rowNumber, "Und 1")),
                    Qty2 = ReadDecimal(Cell(rowNumber, "Qty 2")),
                    Und2 = ReadText(Cell(rowNumber, "Und 2")),

                    // Valores totales
                    FobTotal = ReadDecimal(Cell(rowNumber, "U$ FOB Tot")),
                    FleteTotal = ReadDecimal(Cell(rowNumber, "U$ Flete Tot")),
                    CfrTotal = ReadDecimal(Cell(rowNumber, "U$ CFR Tot")),
                    SeguroTotal = ReadDecimal(Cell(rowNumber, "U$ Seguro Tot")),
                    CifTotal = ReadDecimal(Cell(rowNumber, "U$ CIF Tot")),

                    // Valores unitarios
                    FobUnd1 = ReadDecimal(Cell(rowNumber, "U$ FOB Und 1")),
                    CfrUnd1 = ReadDecimal(Cell(rowNumber, "U$ CFR Und 1")),
                    CifUnd1 = ReadDecimal(Cell(rowNumber, "U$ CIF Und 1")),
                    CifImpUnd1 = ReadDecimal(Cell(rowNumber, "U$ (CIF + Imp) Und 1")),
                    FobUnd2 = ReadDecimal(Cell(rowNumber, "U$ FOB Und 2")),
                    CfrUnd2 = ReadDecimal(Cell(rowNumber, "U$ CFR Und 2")),
                    CifUnd2 = ReadDecimal(Cell(rowNumber, "U$ CIF Und 2")),
                    CifImpUnd2 = ReadDecimal(Cell(rowNumber, "U$ (CIF + Imp) Und 2")),

                    // Procedencia y transporte
                    PaisOrigen = ReadText(Cell(rowNumber, "Pais de Origen")),
                    PaisCompra = ReadText(Cell(rowNumber, "Pais de Compra")),
                    PuertoEmbarque = ReadText(Cell(rowNumber, "Puerto de Embarque")),
                    Via = ReadText(Cell(rowNumber, "Via")),
                    AgenciaTransporte = ReadText(Cell(rowNumber, "Agencia de transporte")),
                    AgenteAduana = ReadText(Cell(rowNumber, "Agente de Aduana")),

                    // Información comercial
                    DescripcionComercial = ReadText(Cell(rowNumber, "Descripcion Comercial")),
                    Igv = ReadDecimal(Cell(rowNumber, "IGV")),
                    Ipm = ReadDecimal(Cell(rowNumber, "IPM")),
                    FechaEmbarque = ReadDate(Cell(rowNumber, "Fecha Embarque")),
                    Descripcion1 = ReadText(Cell(rowNumber, "Descripcion1")),
                    Descripcion2 = ReadText(Cell(rowNumber, "Descripcion2")),
                    Descripcion3 = ReadText(Cell(rowNumber, "Descripcion3")),
                    Descripcion4 = ReadText(Cell(rowNumber, "Descripcion4")),
                    Descripcion5 = ReadText(Cell(rowNumber, "Descripcion5")),
                    Naviera = ReadText(Cell(rowNumber, "Naviera"))
                };

                rows.Add(row);
            }

            return rows;
        }
    }
}
