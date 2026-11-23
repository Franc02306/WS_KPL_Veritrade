
using Sap.Data.Hana;
using Veritrade_Service.Helpers;
using Veritrade_Service.Models;

namespace Veritrade_Service.Services
{
    public class HanaService
    {
        private readonly string _connectionString;
        private readonly ILogger<HanaService> _logger;

        public HanaService(IConfiguration configuration, ILogger<HanaService> logger)
        {
            _connectionString = configuration["Hana:ConnectionString"] ?? throw new Exception("Falta Hana:ConnectionString");
            _logger = logger;
        }

        /// <summary>
        /// Prueba de conexión a la base de datos SAP HANA
        /// </summary>
        public async Task TestConnectionAsync(CancellationToken cancellationToken = default)
        {
            await using var connection = new HanaConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);
            _logger.LogInformation("Conexión a SAP HANA realizada correctamente.");
        }

        /// <summary>
        /// Procesar importadores para un período y partida específicos
        /// </summary>
        public async Task ProcessImportersAsync(int year, int month, string operation, string product, string tariffCode, List<VeritradeDataRow> rows, CancellationToken cancellationToken = default)
        {
            var deleteSql = await SqlQueryLoader.LoadAsync("DeleteImporterPeriod.sql");
            var insertSql = await SqlQueryLoader.LoadAsync("InsertImporter.sql");

            await using var connection = new HanaConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            await using var transaction = connection.BeginTransaction();

            try
            {
                await using (var deleteCommand = new HanaCommand(deleteSql, connection, transaction))
                {
                    deleteCommand.Parameters.Add(new HanaParameter { Value = year });
                    deleteCommand.Parameters.Add(new HanaParameter { Value = month });
                    deleteCommand.Parameters.Add(new HanaParameter { Value = operation });
                    deleteCommand.Parameters.Add(new HanaParameter { Value = product });
                    deleteCommand.Parameters.Add(new HanaParameter { Value = tariffCode });
                    await deleteCommand.ExecuteNonQueryAsync(cancellationToken);
                }

                foreach (var row in rows)
                {
                    await using var command = new HanaCommand(insertSql, connection, transaction);
                    command.Parameters.Add(new HanaParameter { Value = year });
                    command.Parameters.Add(new HanaParameter { Value = month });
                    command.Parameters.Add(new HanaParameter { Value = operation });
                    command.Parameters.Add(new HanaParameter { Value = product });
                    command.Parameters.Add(new HanaParameter { Value = row.Company });
                    command.Parameters.Add(new HanaParameter { Value = row.TotalCifUsd });
                    command.Parameters.Add(new HanaParameter { Value = row.TotalKg });
                    command.Parameters.Add(new HanaParameter { Value = tariffCode });
                    await command.ExecuteNonQueryAsync(cancellationToken);
                }

                transaction.Commit();

                _logger.LogInformation(
                    "Importación confirmada en HANA: {Year}-{Month:00} | {Product} | Partida: {TariffCode} | {Rows} registros",
                    year, month, product, tariffCode, rows.Count);
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }

        /// <summary>
        /// Procesar exportadores para un período y partida específicos
        /// </summary>
        public async Task ProcessExportersAsync(int year, int month, string operation, string product, string tariffCode, List<VeritradeDataRow> rows, CancellationToken cancellationToken = default)
        {
            var deleteSql = await SqlQueryLoader.LoadAsync("DeleteExporterPeriod.sql");
            var insertSql = await SqlQueryLoader.LoadAsync("InsertExporter.sql");

            await using var connection = new HanaConnection(_connectionString);
            await connection.OpenAsync(cancellationToken);

            await using var transaction = connection.BeginTransaction();

            try
            {
                await using (var deleteCommand = new HanaCommand(deleteSql, connection, transaction))
                {
                    deleteCommand.Parameters.Add(new HanaParameter { Value = year });
                    deleteCommand.Parameters.Add(new HanaParameter { Value = month });
                    deleteCommand.Parameters.Add(new HanaParameter { Value = operation });
                    deleteCommand.Parameters.Add(new HanaParameter { Value = product });
                    deleteCommand.Parameters.Add(new HanaParameter { Value = tariffCode });
                    await deleteCommand.ExecuteNonQueryAsync(cancellationToken);
                }

                foreach (var row in rows)
                {
                    await using var command = new HanaCommand(insertSql, connection, transaction);
                    command.Parameters.Add(new HanaParameter { Value = year });
                    command.Parameters.Add(new HanaParameter { Value = month });
                    command.Parameters.Add(new HanaParameter { Value = product });
                    command.Parameters.Add(new HanaParameter { Value = row.Company });
                    command.Parameters.Add(new HanaParameter { Value = row.TotalCifUsd });
                    command.Parameters.Add(new HanaParameter { Value = row.TotalKg });
                    command.Parameters.Add(new HanaParameter { Value = operation });
                    command.Parameters.Add(new HanaParameter { Value = tariffCode });
                    await command.ExecuteNonQueryAsync(cancellationToken);
                }

                transaction.Commit();

                _logger.LogInformation(
                    "Exportación confirmada en HANA: {Year}-{Month:00} | {Product} | Partida: {TariffCode} | {Rows} registros",
                    year, month, product, tariffCode, rows.Count);
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }
    }
}
