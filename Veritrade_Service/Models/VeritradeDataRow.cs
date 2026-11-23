namespace Veritrade_Service.Models
{
    public class VeritradeDataRow
    {
        public string Company { get; set; } = string.Empty;
        public decimal TotalCifUsd { get; set; }
        public decimal TotalKg { get; set; }
    }
}
