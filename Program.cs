using Aepw.Evidence.Listener.Storage;
using Azure.Data.Tables;
using Azure.Identity;
using Azure.Messaging.ServiceBus;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices(services =>
    {
        // Application Insights
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        // =========================================================
        // Microsoft Graph
        // Authenticate as SharePointPOC application in the
        // SharePoint / Graph tenant.
        // =========================================================

        services.AddSingleton(_ =>
            new GraphServiceClient(
                new ClientSecretCredential(
                    GetRequired("GraphTenantId"),
                    GetRequired("GraphClientId"),
                    GetRequired("GraphClientSecret")),
                new[] { "https://graph.microsoft.com/.default" }));


        // =========================================================
        // Azure Managed Identity
        // Used for Service Bus and Table Storage.
        // Registered as a singleton so both clients share one instance.
        // =========================================================

        services.AddSingleton<DefaultAzureCredential>();


        // =========================================================
        // Service Bus
        // Listener only SENDS reference messages.
        // =========================================================

        services.AddSingleton(sp =>
            new ServiceBusClient(
                GetRequired("ServiceBusFullyQualifiedNamespace"),
                sp.GetRequiredService<DefaultAzureCredential>()));

        services.AddSingleton(sp =>
            sp.GetRequiredService<ServiceBusClient>()
              .CreateSender(GetRequired("ServiceBusQueueName")));


        // =========================================================
        // Azure Table Storage
        // Used to store the Microsoft Graph Delta Link.
        // =========================================================

        services.AddSingleton(sp =>
            new TableServiceClient(
                new Uri(GetRequired("StorageAccountTableEndpoint")),
                sp.GetRequiredService<DefaultAzureCredential>()));

        services.AddSingleton<IDeltaStore, TableDeltaStore>();
    })
    .Build();

host.Run();


// =============================================================
// Helper method
// =============================================================

static string GetRequired(string name)
{
    return Environment.GetEnvironmentVariable(name)
        ?? throw new InvalidOperationException(
            $"Missing required app setting '{name}'.");
}