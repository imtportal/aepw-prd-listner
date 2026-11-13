using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace Aepw.Evidence.Listener.Functions;

/// <summary>
/// FUNCTION #3 — subscription renewal (hosted in the Listener app). SharePoint/driveItem change
/// subscriptions expire (driveItem max ≈ 29 days) and cannot be PATCHed once expired. Without
/// renewal, ingestion silently stops. Runs daily and is self-healing: it discovers this Listener's
/// subscription by NotificationUrl, renews it if present, and creates it if not — so it never
/// depends on a stored subscription id being valid.
/// </summary>
public sealed class SubscriptionRenewal
{
    private const int DefaultRenewalDays = 28;
    private const int MaxRenewalDays = 29;   // driveItem subscription cap

    private readonly GraphServiceClient _graph;
    private readonly ILogger<SubscriptionRenewal> _log;

    public SubscriptionRenewal(GraphServiceClient graph, ILogger<SubscriptionRenewal> log)
    {
        _graph = graph;
        _log = log;
    }

    // 02:00 UTC daily. 28-day expiry vs the ~29-day cap leaves a safety margin.
    [Function("SubscriptionRenewal")]
    public async Task Run([TimerTrigger("0 0 2 * * *")] TimerInfo timer)
    {
        var notificationUrl = Environment.GetEnvironmentVariable("NotificationUrl");
        var clientState = Environment.GetEnvironmentVariable("WebhookClientState");
        var driveId = Environment.GetEnvironmentVariable("GraphDriveId");
        var resource = Environment.GetEnvironmentVariable("GraphResource")
                       ?? (string.IsNullOrWhiteSpace(driveId) ? null : $"drives/{driveId}/root");

        var missing = new List<string>();
        if (string.IsNullOrWhiteSpace(notificationUrl)) missing.Add("NotificationUrl");
        if (string.IsNullOrWhiteSpace(resource)) missing.Add("GraphResource (or GraphDriveId)");
        if (string.IsNullOrWhiteSpace(clientState)) missing.Add("WebhookClientState");
        if (missing.Count > 0)
        {
            _log.LogError("Subscription renewal cannot run; missing app setting(s): {Missing}.",
                string.Join(", ", missing));
            return;
        }

        var renewalDays = DefaultRenewalDays;
        if (int.TryParse(Environment.GetEnvironmentVariable("RenewalDays"), out var configured) && configured > 0)
            renewalDays = configured;
        renewalDays = Math.Min(renewalDays, MaxRenewalDays);

        var newExpiry = DateTimeOffset.UtcNow.AddDays(renewalDays);

        try
        {
            // Graph returns only subscriptions owned by the calling identity.
            var subs = await _graph.Subscriptions.GetAsync();

            var existing = subs?.Value?.FirstOrDefault(s =>
                string.Equals(s.NotificationUrl, notificationUrl, StringComparison.OrdinalIgnoreCase)
                && string.Equals(Normalize(s.Resource), Normalize(resource), StringComparison.OrdinalIgnoreCase));

            if (existing?.Id is not null)
            {
                await _graph.Subscriptions[existing.Id].PatchAsync(
                    new Subscription { ExpirationDateTime = newExpiry });

                _log.LogInformation("Renewed subscription {Sub} to {Expiry:o}.", existing.Id, newExpiry);
                return;
            }

            var created = await _graph.Subscriptions.PostAsync(new Subscription
            {
                ChangeType         = "updated",
                NotificationUrl    = notificationUrl,
                Resource           = resource,
                ClientState        = clientState,
                ExpirationDateTime = newExpiry
            });

            _log.LogInformation("Created subscription {Sub} for {Resource}, expires {Expiry:o}.",
                created?.Id, resource, newExpiry);
        }
        catch (ODataError ex)
        {
            // Surface as an error so an Application Insights alert can page on-call: a persistent
            // failure here means change notifications will stop arriving.
            _log.LogError(ex,
                "Subscription renew/create FAILED. Status={Status} Code={Code} Message={Message}",
                ex.ResponseStatusCode, ex.Error?.Code, ex.Error?.Message);
            throw;
        }
    }

    // Graph may echo the resource with/without a leading slash; compare on a normalized form.
    private static string Normalize(string? resource) => (resource ?? string.Empty).Trim().Trim('/');
}
