namespace VeritradeDownloader.Models
{
    public class VeritradeImporter
    {
        public string Importer { get; set; } = string.Empty;
        public int TotalRecords { get; set; }
        public decimal TotalCifUsd { get; set; }
        public decimal TotalKg { get; set; }
        public decimal UsdPerKg { get; set; }
    }
}
