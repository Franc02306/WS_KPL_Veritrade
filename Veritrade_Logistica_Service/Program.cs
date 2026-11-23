
using Azure.Identity;
using Microsoft.Graph;
using Veritrade_Logistica_Service;
using Veritrade_Logistica_Service.Services;

var builder = Host.CreateApplicationBuilder(args);

builder.Services.AddWindowsService(options =>
{
    options.ServiceName = "Veritrade_Logistica_Service";
});

// Configuración de Microsoft Graph
var tenantId = builder.Configuration["AzureAd:TenantId"]
    ?? throw new Exception("Falta AzureAd:TenantId");

var clientId = builder.Configuration["AzureAd:ClientId"]
    ?? throw new Exception("Falta AzureAd:ClientId");

var clientSecret = builder.Configuration["AzureAd:ClientSecret"]
    ?? throw new Exception("Falta AzureAd:ClientSecret");

var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

var graphClient = new GraphServiceClient(
    credential,
    new[]
    {
        "https://graph.microsoft.com/.default"
    }
);

// Registrar dependencias
builder.Services.AddSingleton(graphClient);
builder.Services.AddSingleton<OneDriveFileService>();
//builder.Services.AddSingleton<HanaService>();

// Servicio principal
builder.Services.AddHostedService<Worker>();

var host = builder.Build();

host.Run();
