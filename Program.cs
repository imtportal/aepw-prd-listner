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

        var graphTenantId = GetRequired("GraphTenantId");
        var graphClientId = GetRequired("GraphClientId");
        var graphClientSecret = GetRequired("GraphClientSecret");

        var graphCredential = new ClientSecretCredential(
            graphTenantId,
            graphClientId,
            graphClientSecret);

        services.AddSingleton(_ =>
            new GraphServiceClient(
                graphCredential,
                new[] { "https://graph.microsoft.com/.default" }));


        // =========================================================
        // Azure Managed Identity
        // Used for Service Bus and Table Storage.
        // =========================================================

        var managedIdentityCredential = new DefaultAzureCredential();


        // =========================================================
        // Service Bus
        // Listener only SENDS reference messages.
        // =========================================================

        services.AddSingleton(_ =>
        {
            var namespaceName =
                GetRequired("ServiceBusFullyQualifiedNamespace");

            return new ServiceBusClient(
                namespaceName,
                managedIdentityCredential);
        });

        services.AddSingleton(sp =>
        {
            var client =
                sp.GetRequiredService<ServiceBusClient>();

            var queue =
                GetRequired("ServiceBusQueueName");

            return client.CreateSender(queue);
        });


        // =========================================================
        // Azure Table Storage
        // Used to store the Microsoft Graph Delta Link.
        // =========================================================

        services.AddSingleton(_ =>
        {
            var endpoint =
                new Uri(GetRequired("StorageAccountTableEndpoint"));

            return new TableServiceClient(
                endpoint,
                managedIdentityCredential);
        });

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