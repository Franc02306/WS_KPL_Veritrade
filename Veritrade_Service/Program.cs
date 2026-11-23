using Veritrade_Service.Services;
using Azure.Identity;
using Microsoft.Graph;
using Veritrade_Service;
using Veritrade_Service.Logging;

var builder = Host.CreateApplicationBuilder(args);

builder.Services.AddWindowsService(options =>
{
    options.ServiceName = "Veritrade_Service";
});

var tenantId = builder.Configuration["AzureAd:TenantId"] ?? throw new Exception("Falta AzureAd:TenantId");

var clientId = builder.Configuration["AzureAd:ClientId"] ?? throw new Exception("Falta AzureAd:ClientId");

var clientSecret = builder.Configuration["AzureAd:ClientSecret"] ?? throw new Exception("Falta AzureAd:ClientSecret");

var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

var graphClient = new GraphServiceClient(credential, new[]
    {
        "https://graph.microsoft.com/.default"
    }
);

builder.Services.AddSingleton(graphClient);

builder.Services.AddSingleton<OneDriveFileService>();
builder.Services.AddSingleton<HanaService>();
builder.Services.AddHostedService<Worker>();

// Configuración de Logs
builder.Logging.ClearProviders();
builder.Logging.AddProvider(new FileLoggerProvider());

var host = builder.Build();

host.Run();