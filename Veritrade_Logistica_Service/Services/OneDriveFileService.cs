using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Text;

namespace Veritrade_Logistica_Service.Services
{
    public class OneDriveFileService
    {
        private readonly GraphServiceClient _graphClient;
        private readonly IConfiguration _configuration;
        private readonly ILogger<OneDriveFileService> _logger;

        public OneDriveFileService(GraphServiceClient graphClient, IConfiguration configuration, ILogger<OneDriveFileService> logger)
        {
            _graphClient = graphClient;
            _configuration = configuration;
            _logger = logger;
        }

        /// <summary>
        /// Mover un archivo a la carpeta "Processed" en OneDrive
        /// </summary>
        public async Task MoveToProcessedAsync(string driveId, string itemId, string fileName, string relativePath, CancellationToken cancellationToken)
        {
            var url = _configuration["Veritrade:OneDrive:ProcessedUrl"] ?? throw new Exception("Falta Veritrade:OneDrive:ProcessedUrl");
            await MoveAsync(driveId, itemId, fileName, relativePath, url, cancellationToken);
            _logger.LogInformation("{FileName} movido a Processed/{Path}", fileName, relativePath);
        }

        /// <summary>
        /// Mover un archivo a la carpeta "Error" en OneDrive
        /// </summary>
        public async Task MoveToErrorAsync(string driveId, string itemId, string fileName, string relativePath, CancellationToken cancellationToken)
        {
            var url = _configuration["Veritrade:OneDrive:ErrorUrl"] ?? throw new Exception("Falta Veritrade:OneDrive:ErrorUrl");
            await MoveAsync(driveId, itemId, fileName, relativePath, url, cancellationToken);
            _logger.LogWarning("{FileName} movido a Error/{Path}", fileName, relativePath);
        }

        /// <summary>
        /// Mover un archivo a una carpeta compartida en OneDrive
        /// </summary>
        private async Task MoveAsync(string driveId, string itemId, string fileName, string relativePath, string destinationUrl, CancellationToken cancellationToken)
        {
            var destinationRoot = await GetSharedFolderAsync(destinationUrl, cancellationToken);

            var destinationDriveId = destinationRoot.ParentReference?.DriveId ?? throw new Exception("No se pudo obtener el DriveId destino");

            if (destinationDriveId != driveId)
                throw new Exception("Input y destino deben pertenecer al mismo OneDrive.");

            var destinationFolderId = destinationRoot.Id ?? throw new Exception("No se pudo obtener la carpeta destino.");

            foreach (var folderName in relativePath.Split('/', StringSplitOptions.RemoveEmptyEntries))
                destinationFolderId = await GetOrCreateFolderAsync(driveId, destinationFolderId, folderName, cancellationToken);

            var children = await _graphClient.Drives[driveId].Items[destinationFolderId].Children.GetAsync(cancellationToken: cancellationToken);

            var existing = children?.Value?.FirstOrDefault(x =>
                x.File is not null &&
                x.Name != null &&
                x.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase));

            if (existing?.Id is not null)
                await _graphClient.Drives[driveId].Items[existing.Id].DeleteAsync(cancellationToken: cancellationToken);

            var request = new DriveItem
            {
                Name = fileName,
                ParentReference = new ItemReference
                {
                    Id = destinationFolderId
                }
            };

            await _graphClient.Drives[driveId].Items[itemId].PatchAsync(request, cancellationToken: cancellationToken);
        }

        /// <summary>
        /// Obtener o crear una carpeta en OneDrive
        /// Si la carpeta ya existe, devuelve su ID
        /// si no, la crea y devuelve el ID de la nueva carpeta
        /// </summary>
        private async Task<string> GetOrCreateFolderAsync(string driveId, string parentId, string folderName, CancellationToken cancellationToken)
        {
            var children = await _graphClient.Drives[driveId].Items[parentId].Children.GetAsync(cancellationToken: cancellationToken);

            var existing = children?.Value?.FirstOrDefault(x =>
                x.Folder is not null &&
                x.Name != null &&
                x.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase));

            if (existing?.Id is not null)
                return existing.Id;

            var folder = new DriveItem
            {
                Name = folderName,
                Folder = new Folder()
            };

            var created = await _graphClient.Drives[driveId].Items[parentId].Children.PostAsync(folder, cancellationToken: cancellationToken);

            return created?.Id ?? throw new Exception($"No se pudo crear la carpeta {folderName}");
        }

        /// <summary>
        /// Obtener la carpeta compartida a partir de la URL de compartición
        /// </summary>
        private async Task<DriveItem> GetSharedFolderAsync(string url, CancellationToken cancellationToken)
        {
            var shareId = EncodeSharingUrl(url);

            return await _graphClient.Shares[shareId].DriveItem.GetAsync(cancellationToken: cancellationToken)
                ?? throw new Exception("No se pudo obtener la carpeta compartida.");
        }

        /// <summary>
        /// Obtener el ID de compartición a partir de la URL de compartición
        /// </summary>
        private static string EncodeSharingUrl(string url)
        {
            var bytes = Encoding.UTF8.GetBytes(url);
            return "u!" + Convert.ToBase64String(bytes).TrimEnd('=').Replace('/', '_').Replace('+', '-');
        }
    }
}
