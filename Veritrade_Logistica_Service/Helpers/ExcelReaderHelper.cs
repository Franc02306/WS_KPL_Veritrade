using ClosedXML.Excel;
using System.Globalization;

namespace Veritrade_Logistica_Service.Helpers
{
    public static class ExcelReaderHelper
    {
        /// <summary>
        /// Leer valores de texto del Excel
        /// </summary>
        public static string? ReadText(IXLCell cell)
        {
            if (cell.IsEmpty())
                return null;

            var value = cell.GetString().Replace("\0", "").Replace("_x0000_", "").Trim();
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }

        /// <summary>
        /// Leer valores numéricos del Excel
        /// </summary>
        public static decimal? ReadDecimal(IXLCell cell)
        {
            if (cell.IsEmpty())
                return null;

            if (cell.TryGetValue<decimal>(out var number))
                return number;

            var value = cell.GetString().Trim();

            if (decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out number))
                return number;

            if (decimal.TryParse(value, NumberStyles.Number, CultureInfo.GetCultureInfo("es-PE"), out number))
                return number;

            throw new Exception($"Valor numérico inválido en {cell.Address}: {value}");
        }

        /// <summary>
        /// Leer fechas del Excel
        /// </summary>
        public static DateTime? ReadDate(IXLCell cell)
        {
            if (cell.IsEmpty())
                return null;

            if (cell.TryGetValue<DateTime>(out var date))
                return date.Year is >= 2000 and <= 2100 ? date.Date : null;

            if (cell.TryGetValue<double>(out var serial) && serial >= new DateTime(2000, 1, 1).ToOADate() && serial < new DateTime(2101, 1, 1).ToOADate())
                return DateTime.FromOADate(serial).Date;

            if (DateTime.TryParse(cell.GetString().Trim(), CultureInfo.GetCultureInfo("es-PE"), DateTimeStyles.None, out date))
                return date.Year is >= 2000 and <= 2100 ? date.Date : null;

            return null;
        }
    }
}
