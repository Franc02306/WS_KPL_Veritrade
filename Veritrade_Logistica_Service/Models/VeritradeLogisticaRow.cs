namespace Veritrade_Logistica_Service.Models
{
    public class VeritradeLogisticaRow
    {
        // Control del servicio
        public int Anio { get; set; }
        public string Producto { get; set; } = string.Empty;
        public string Operacion { get; set; } = "I";
        public int FilaExcel { get; set; }

        // Información aduanera
        public string? PartidaAduanera { get; set; }
        public string? DescripcionPartida { get; set; }
        public string? Aduana { get; set; }
        public string? DuaDam { get; set; }
        public DateTime? Fecha { get; set; }
        public DateTime? Eta { get; set; }
        public string? CodTributario { get; set; }
        public string? Importador { get; set; }
        public string? EmbarcadorExportador { get; set; }

        // Pesos y cantidades
        public decimal? KgBruto { get; set; }
        public decimal? KgNeto { get; set; }
        public decimal? Qty1 { get; set; }
        public string? Und1 { get; set; }
        public decimal? Qty2 { get; set; }
        public string? Und2 { get; set; }

        // Valores totales
        public decimal? FobTotal { get; set; }
        public decimal? FleteTotal { get; set; }
        public decimal? CfrTotal { get; set; }
        public decimal? SeguroTotal { get; set; }
        public decimal? CifTotal { get; set; }

        // Valores unitarios
        public decimal? FobUnd1 { get; set; }
        public decimal? CfrUnd1 { get; set; }
        public decimal? CifUnd1 { get; set; }
        public decimal? CifImpUnd1 { get; set; }
        public decimal? FobUnd2 { get; set; }
        public decimal? CfrUnd2 { get; set; }
        public decimal? CifUnd2 { get; set; }
        public decimal? CifImpUnd2 { get; set; }

        // Procedencia y transporte
        public string? PaisOrigen { get; set; }
        public string? PaisCompra { get; set; }
        public string? PuertoEmbarque { get; set; }
        public string? Via { get; set; }
        public string? AgenciaTransporte { get; set; }
        public string? AgenteAduana { get; set; }

        // Información comercial
        public string? DescripcionComercial { get; set; }
        public decimal? Igv { get; set; }
        public decimal? Ipm { get; set; }
        public DateTime? FechaEmbarque { get; set; }
        public string? Descripcion1 { get; set; }
        public string? Descripcion2 { get; set; }
        public string? Descripcion3 { get; set; }
        public string? Descripcion4 { get; set; }
        public string? Descripcion5 { get; set; }
        public string? Naviera { get; set; }
    }
}
