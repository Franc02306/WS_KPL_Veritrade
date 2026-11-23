using ClosedXML.Excel;
using VeritradeDownloader.Models;

namespace VeritradeDownloader.Services
{
    public class ExcelReaderService
    {
        public void Read(string filePath)
        {
            using var workbook = new XLWorkbook(filePath);

            Console.WriteLine($"Archivo: {Path.GetFileName(filePath)}");

            foreach (var worksheet in workbook.Worksheets)
            {
                Console.WriteLine($"Hoja encontrada: {worksheet.Name}");
            }

            ReadImporters(workbook);
        }

        private void ReadImporters(XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheet("Importadores");

            var countryOperation = worksheet.Cell("B1").GetString().Trim();
            var filter = worksheet.Cell("B2").GetString().Trim();
            var period = worksheet.Cell("B3").GetString().Trim();

            Console.WriteLine();
            Console.WriteLine($"Reporte: {countryOperation}");
            Console.WriteLine($"Filtro: {filter}");
            Console.WriteLine($"Periodo: {period}");
            Console.WriteLine();

            var row = 7;

            while (true)
            {
                var importerName = worksheet.Cell(row, 1).GetString().Trim();

                if (string.IsNullOrWhiteSpace(importerName) ||
                    importerName.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                var importer = new VeritradeImporter
                {
                    Importer = importerName,
                    TotalRecords = worksheet.Cell(row, 2).GetValue<int>(),
                    TotalCifUsd = worksheet.Cell(row, 3).GetValue<decimal>(),
                    TotalKg = worksheet.Cell(row, 5).GetValue<decimal>(),
                    UsdPerKg = worksheet.Cell(row, 6).GetValue<decimal>()
                };

                Console.WriteLine(
                    $"{importer.Importer} | " +
                    $"{importer.TotalRecords} | " +
                    $"{importer.TotalCifUsd:N2} | " +
                    $"{importer.TotalKg:N2}"
                );

                row++;
            }
        }
    }
}
