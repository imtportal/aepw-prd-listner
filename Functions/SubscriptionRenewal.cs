using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace Aepw.Evidence.Listener.Functions;

/// <summary>
/// FUNCTION #3 — subscription renewal (hosted in the Listener app). SharePoint/driveItem change
/// subscriptions expire (driveItem max ≈ 30 days). Without renewal, ingestion silently stops.
/// Runs daily and pushes the expiry back to a safe margin inside the cap.
/// </summary>
public sealed class SubscriptionRenewal
{
    private readonly GraphServiceClient _graph;
    private readonly ILogger<SubscriptionRenewal> _log;

    public SubscriptionRenewal(GraphServiceClient graph, ILogger<SubscriptionRenewal> log)
    {
        _graph = graph;
        _log = log;
    }

    // 02:00 UTC daily. 28-day expiry vs the ~30-day cap leaves a wide safety margin.
    [Function("SubscriptionRenewal")]
    public async Task Run([TimerTrigger("0 0 2 * * *")] TimerInfo timer)
    {
        var subscriptionId = Environment.GetEnvironmentVariable("SubscriptionId");
        if (string.IsNullOrEmpty(subscriptionId))
        {
            _log.LogWarning("SubscriptionId not configured; nothing to renew.");
            return;
        }

        try
        {
            var patch = new Subscription
            {
                ExpirationDateTime = DateTimeOffset.UtcNow.AddDays(28)
            };

            await _graph.Subscriptions[subscriptionId].PatchAsync(patch);

            _log.LogInformation("Renewed subscription {Sub} to {Expiry:o}.",
                subscriptionId, patch.ExpirationDateTime);
        }
        catch (Exception ex)
        {
            // Surface as an error so an Application Insights alert can page on-call: a persistent
            // failure here means change notifications will stop arriving.
            _log.LogError(ex, "Subscription renewal FAILED for {Sub}.", subscriptionId);
            throw;
        }
    }
}
