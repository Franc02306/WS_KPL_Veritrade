using Sap.Data.Hana;
using Veritrade_Logistica_Service.Helpers;
using Veritrade_Logistica_Service.Models;

namespace Veritrade_Logistica_Service.Services
{
    public class HanaService
    {
        private readonly string _connectionString;
        private readonly ILogger<HanaService> _logger;

        private static readonly (string Column, Func<VeritradeLogisticaRow, object?> Value)[] Columns =
        {
            ("ANIO", x => x.Anio),
            ("PRODUCTO", x => x.Producto),
            ("OPERACION", x => x.Operacion),
            ("FILA_EXCEL", x => x.FilaExcel),

            // Información aduanera
            ("PARTIDA_ADUANERA", x => x.PartidaAduanera),
            ("DESCRIPCION_PARTIDA", x => x.DescripcionPartida),
            ("ADUANA", x => x.Aduana),
            ("DUA_DAM", x => x.DuaDam),
            ("FECHA", x => x.Fecha),
            ("ETA", x => x.Eta),
            ("COD_TRIBUTARIO", x => x.CodTributario),
            ("IMPORTADOR", x => x.Importador),
            ("EMBARCADOR_EXPORTADOR", x => x.EmbarcadorExportador),

            // Pesos y cantidades
            ("KG_BRUTO", x => x.KgBruto),
            ("KG_NETO", x => x.KgNeto),
            ("QTY_1", x => x.Qty1),
            ("UND_1", x => x.Und1),
            ("QTY_2", x => x.Qty2),
            ("UND_2", x => x.Und2),

            // Valores totales
            ("FOB_TOTAL", x => x.FobTotal),
            ("FLETE_TOTAL", x => x.FleteTotal),
            ("CFR_TOTAL", x => x.CfrTotal),
            ("SEGURO_TOTAL", x => x.SeguroTotal),
            ("CIF_TOTAL", x => x.CifTotal),

            // Valores unitarios
            ("FOB_UND_1", x => x.FobUnd1),
            ("CFR_UND_1", x => x.CfrUnd1),
            ("CIF_UND_1", x => x.CifUnd1),
            ("CIF_IMP_UND_1", x => x.CifImpUnd1),
            ("FOB_UND_2", x => x.FobUnd2),
            ("CFR_UND_2", x => x.CfrUnd2),
            ("CIF_UND_2", x => x.CifUnd2),
            ("CIF_IMP_UND_2", x => x.CifImpUnd2),

            // Procedencia y transporte
            ("PAIS_ORIGEN", x => x.PaisOrigen),
            ("PAIS_COMPRA", x => x.PaisCompra),
            ("PUERTO_EMBARQUE", x => x.PuertoEmbarque),
            ("VIA", x => x.Via),
            ("AGENCIA_TRANSPORTE", x => x.AgenciaTransporte),
            ("AGENTE_ADUANA", x => x.AgenteAduana),

            // Información comercial
            ("DESCRIPCION_COMERCIAL", x => x.DescripcionComercial),
            ("IGV", x => x.Igv),
            ("IPM", x => x.Ipm),
            ("FECHA_EMBARQUE", x => x.FechaEmbarque),
            ("DESCRIPCION1", x => x.Descripcion1),
            ("DESCRIPCION2", x => x.Descripcion2),
            ("DESCRIPCION3", x => x.Descripcion3),
            ("DESCRIPCION4", x => x.Descripcion4),
            ("DESCRIPCION5", x => x.Descripcion5),
            ("NAVIERA", x => x.Naviera)
        };

        public HanaService(IConfiguration configuration, ILogger<HanaService> logger)
        {
            _connectionString = configuration["Hana:ConnectionString"] ?? throw new Exception("Falta Hana:ConnectionString");
            _logger = logger;
        }

        /// <summary>
        /// Prueba de conexión a SAP HANA
        /// </summary>
        public async Task TestConnectionAsync(CancellationToken cancellationToken = default)
        {
            await using var connection = new HanaConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);
            _logger.LogInformation("Conexión a SAP HANA Logistica realizada correctamente.");
        }

        /// <summary>
        /// Procesar importaciones anuales por producto y partida
        /// </summary>
        public async Task ProcessImportersAsync(int year, string product, string tariffCode, List<VeritradeLogisticaRow> rows, CancellationToken cancellationToken = default)
        {
            if (rows is null || rows.Count == 0)
                throw new Exception("El Excel no contiene registros válidos. No se modificará HANA.");

            product = product.Trim().ToUpperInvariant();
            tariffCode = tariffCode.Trim();

            if (year < 2000 || year > 9999 || string.IsNullOrWhiteSpace(product) || string.IsNullOrWhiteSpace(tariffCode))
                throw new ArgumentException("Año, producto o partida inválidos.");

            if (rows.Any(x => x.Anio != year || !string.Equals(x.Producto, product, StringComparison.OrdinalIgnoreCase) || x.Operacion != "I" || x.PartidaAduanera?.Trim() != tariffCode))
                throw new Exception("Los registros no coinciden con el año, producto, operación o partida del archivo.");

            var deleteSql = await SqlQueryLoader.LoadAsync("DeleteImporterYear.sql", cancellationToken);
            var insertSql = await SqlQueryLoader.LoadAsync("InsertImporter.sql", cancellationToken);

            await using var connection = new HanaConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            await using var transaction = connection.BeginTransaction();

            try
            {
                await using (var deleteCommand = new HanaCommand(deleteSql, connection, transaction))
                {
                    deleteCommand.Parameters.Add(new HanaParameter { Value = year });
                    deleteCommand.Parameters.Add(new HanaParameter { Value = product });
                    deleteCommand.Parameters.Add(new HanaParameter { Value = "I" });
                    deleteCommand.Parameters.Add(new HanaParameter { Value = tariffCode });
                    await deleteCommand.ExecuteNonQueryAsync(cancellationToken);
                }

                foreach (var row in rows)
                {
                    await using var command = new HanaCommand(insertSql, connection, transaction);

                    foreach (var column in Columns)
                        command.Parameters.Add(new HanaParameter { Value = column.Value(row) ?? DBNull.Value });

                    await command.ExecuteNonQueryAsync(cancellationToken);
                }

                transaction.Commit();

                _logger.LogInformation("Importación anual confirmada: {Year} | {Product} | Partida: {TariffCode} | {Rows} registros", year, product, tariffCode, rows.Count);
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }
    }
}
