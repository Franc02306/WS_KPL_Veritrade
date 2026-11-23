namespace Veritrade_Service.Models
{
    public class VeritradeReportConfig
    {
        public string Product { get; set; } = string.Empty;
        public string Operation { get; set; } = string.Empty;
        public List<string> TariffCodes { get; set; } = [];
    }
}
