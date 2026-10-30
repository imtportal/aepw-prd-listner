using System.Net;
using System.Text.Json;
using Aepw.Evidence.Listener.Models;
using Aepw.Evidence.Listener.Storage;
using Azure.Messaging.ServiceBus;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace Aepw.Evidence.Listener.Functions;

/// <summary>
/// FUNCTION #1 — the webhook listener. Two jobs only:
///   (1) answer the one-time Graph validation handshake, and
///   (2) on a change notification, ACK within ~5s, resolve exactly which items changed via a
///       Graph delta query, and enqueue a lightweight reference for each.
/// The file is NOT downloaded here — that is the Worker's job. This keeps the ACK fast and the
/// queue message tiny.
///
/// NOTE: AuthorizationLevel.Anonymous — a public webhook does not use a function key. Authenticity
/// is enforced by the Graph validation handshake plus the clientState secret checked below.
/// </summary>
public sealed class SharePointWebhookListener
{
    private readonly ServiceBusSender _sender;
    private readonly GraphServiceClient _graph;
    private readonly IDeltaStore _deltaStore;
    private readonly ILogger<SharePointWebhookListener> _log;

    public SharePointWebhookListener(
        ServiceBusSender sender,
        GraphServiceClient graph,
        IDeltaStore deltaStore,
        ILogger<SharePointWebhookListener> log)
    {
        _sender = sender;
        _graph = graph;
        _deltaStore = deltaStore;
        _log = log;
    }

    [Function("SharePointWebhook")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post")] HttpRequestData req)
    {
        // ---- 1) VALIDATION HANDSHAKE -------------------------------------------------------
        // On subscription creation Graph calls POST {url}?validationToken=... and expects the
        // token echoed back verbatim as text/plain within ~5 seconds.

        // Diagnostic logging — safe to remove once the subscription is registered and stable.
        _log.LogInformation("Request URL: {Url}", req.Url.ToString());
        _log.LogInformation("Raw Query: {Query}", req.Url.Query);

        var validationToken = GetQueryValue(req.Url.Query, "validationToken");
        _log.LogInformation("Extracted validation token: [{Token}]", validationToken);

        if (!string.IsNullOrEmpty(validationToken))
        {
            _log.LogInformation("Validation handshake received.");

            var ok = req.CreateResponse(HttpStatusCode.OK);
            ok.Headers.Add("Content-Type", "text/plain; charset=utf-8");
            await ok.WriteStringAsync(validationToken);

            return ok;
        }

        // ---- 2) CHANGE NOTIFICATION --------------------------------------------------------
        var body = await new StreamReader(req.Body).ReadToEndAsync();
        _log.LogInformation("Webhook notification received. BodyLength={BodyLength}", body.Length);

        var expectedClientState = Environment.GetEnvironmentVariable("WebhookClientState");
        var driveId = Environment.GetEnvironmentVariable("GraphDriveId")
                      ?? throw new InvalidOperationException("Missing app setting 'GraphDriveId'.");
        var siteId = Environment.GetEnvironmentVariable("GraphSiteId") ?? string.Empty;

        try
        {
            var envelope = string.IsNullOrWhiteSpace(body)
                ? null
                : JsonSerializer.Deserialize<GraphNotificationEnvelope>(body);

            var notifications = envelope?.Value ?? new List<GraphNotification>();
            _log.LogInformation(
                "Parsed webhook notification. NotificationCount={NotificationCount}, DriveId={DriveId}, SiteIdConfigured={SiteIdConfigured}",
                notifications.Count,
                driveId,
                !string.IsNullOrWhiteSpace(siteId));

            foreach (var n in notifications)
            {
                _log.LogInformation(
                    "Notification received. SubscriptionId={SubscriptionId}, Resource={Resource}, ChangeType={ChangeType}",
                    n.SubscriptionId,
                    n.Resource,
                    n.ChangeType);

                // Reject spoofed callers: clientState must match the secret set at subscribe time.
                if (!string.IsNullOrEmpty(expectedClientState) && n.ClientState != expectedClientState)
                {
                    _log.LogWarning("clientState mismatch — ignoring notification {Sub}.", n.SubscriptionId);
                    continue;
                }

                _log.LogInformation(
                    "Starting delta processing for notification. SubscriptionId={SubscriptionId}, DriveId={DriveId}",
                    n.SubscriptionId,
                    driveId);

                await ProcessDeltaAsync(siteId, driveId);
            }
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Delta processing failed");

            _log.LogError("Exception Type: {Type}", ex.GetType().FullName);
            _log.LogError("Exception Message: {Message}", ex.Message);

            if (ex.InnerException != null)
            {
                _log.LogError("Inner Exception: {Inner}", ex.InnerException.ToString());
            }

            throw;
        }

        return req.CreateResponse(HttpStatusCode.Accepted);   // 202
    }

    /// <summary>
    /// Resolves which items changed using the stored delta link and enqueues a reference for each.
    /// Handles nextLink paging within a single run and persists the fresh deltaLink at the end.
    /// </summary>
    private async Task ProcessDeltaAsync(string siteId, string driveId)
    {
        _log.LogInformation("DELTA START | DriveId={DriveId}", driveId);

        var deltaLink = await _deltaStore.GetDeltaLinkAsync(driveId);

        _log.LogInformation(
            "DELTA LINK | DriveId={DriveId} | HasStoredDeltaLink={HasStoredDeltaLink}",
            driveId,
            !string.IsNullOrWhiteSpace(deltaLink));

        // First run: enumerate from the drive's real root item id.
        // Subsequent runs: continue from the stored deltaLink. WithUrl(...) overrides the request
        // URL entirely, so Items["root"] there is just a placeholder used to obtain the builder.
        // Both branches call Items[..].Delta.GetAsDeltaGetResponseAsync() so `page` is one type.
        var rootId = deltaLink is null ? (await _graph.Drives[driveId].Root.GetAsync())?.Id : null;

        var page = deltaLink is null
            ? await _graph.Drives[driveId].Items[rootId!].Delta.GetAsDeltaGetResponseAsync()
            : await _graph.Drives[driveId].Items["root"].Delta
                  .WithUrl(deltaLink).GetAsDeltaGetResponseAsync();

        int enqueued = 0;
        int deltaItemsSeen = 0;
        string? freshDeltaLink = null;

        while (page is not null)
        {
            _log.LogInformation(
                "DELTA PAGE | DriveId={DriveId} | ItemCount={ItemCount} | HasNextLink={HasNextLink} | HasDeltaLink={HasDeltaLink}",
                driveId,
                page.Value?.Count ?? 0,
                !string.IsNullOrWhiteSpace(page.OdataNextLink),
                !string.IsNullOrWhiteSpace(page.OdataDeltaLink));

            foreach (var item in page.Value ?? new List<DriveItem>())
            {
                deltaItemsSeen++;

                var isDelete = item.Deleted is not null;

                _log.LogInformation(
                    "DELTA ITEM | Name={Name} | ItemId={ItemId} | IsFile={IsFile} | IsFolder={IsFolder} | IsRoot={IsRoot} | IsDeleted={IsDeleted} | ETag={ETag} | Size={Size}",
                    item.Name,
                    item.Id,
                    item.File is not null,
                    item.Folder is not null,
                    item.Root is not null,
                    isDelete,
                    item.ETag,
                    item.Size);

                // Skip folders and the root container itself; act only on files.
                if (item.Folder is not null || item.Root is not null)
                {
                    _log.LogInformation(
                        "DELTA ITEM SKIPPED | Name={Name} | ItemId={ItemId} | Reason={Reason}",
                        item.Name,
                        item.Id,
                        item.Folder is not null ? "Folder" : "Root container");

                    continue;
                }

                var msg = new ProcessingRequest
                {
                    EventType = isDelete ? "deleted" : "created_or_updated",
                    SiteId    = siteId,
                    DriveId   = driveId,
                    ItemId    = item.Id ?? string.Empty,
                    ETag      = item.ETag,
                    FileName  = item.Name ?? string.Empty,
                    SizeBytes = item.Size ?? 0,
                    WebUrl    = item.WebUrl
                };

                var payload = JsonSerializer.Serialize(msg);

                var sb = new ServiceBusMessage(payload)
                {
                    MessageId     = msg.DedupKey,   // native dedup — enable duplicate detection on the queue
                    CorrelationId = msg.CorrelationId,
                    ContentType   = "application/json"
                };

                _log.LogInformation(
                    "SB SEND START | FileName={FileName} | ItemId={ItemId} | ETag={ETag} | MessageId={MessageId} | CorrelationId={CorrelationId} | EventType={EventType} | BodyLength={BodyLength} | Body={Body}",
                    msg.FileName,
                    msg.ItemId,
                    msg.ETag,
                    sb.MessageId,
                    sb.CorrelationId,
                    msg.EventType,
                    sb.Body.ToMemory().Length,
                    payload);

                await _sender.SendMessageAsync(sb);

                enqueued++;

                _log.LogInformation(
                    "SB SEND SUCCESS | FileName={FileName} | ItemId={ItemId} | MessageId={MessageId} | EnqueuedCount={EnqueuedCount}",
                    msg.FileName,
                    msg.ItemId,
                    sb.MessageId,
                    enqueued);
            }

            // The deltaLink appears only on the final page; follow nextLink paging until then.
            if (!string.IsNullOrEmpty(page.OdataDeltaLink))
                freshDeltaLink = page.OdataDeltaLink;

            if (!string.IsNullOrEmpty(page.OdataNextLink))
            {
                _log.LogInformation("DELTA NEXT PAGE | DriveId={DriveId}", driveId);

                page = await _graph.Drives[driveId].Items["root"].Delta
                    .WithUrl(page.OdataNextLink).GetAsDeltaGetResponseAsync();
            }
            else
            {
                break;
            }
        }

        _log.LogInformation(
            "DELTA COMPLETE | DriveId={DriveId} | DeltaItemsSeen={DeltaItemsSeen} | Enqueued={Enqueued} | HasFreshDeltaLink={HasFreshDeltaLink}",
            driveId,
            deltaItemsSeen,
            enqueued,
            !string.IsNullOrWhiteSpace(freshDeltaLink));

        await _deltaStore.SaveDeltaLinkAsync(driveId, freshDeltaLink);

        _log.LogInformation(
            "DELTA LINK SAVED | DriveId={DriveId} | HasFreshDeltaLink={HasFreshDeltaLink}",
            driveId,
            !string.IsNullOrWhiteSpace(freshDeltaLink));

        _log.LogInformation("Delta enqueued {Count} item(s) for drive {Drive}.", enqueued, driveId);
    }

    /// <summary>
    /// Minimal, dependency-free query-string reader (case-insensitive). Avoids taking a dependency
    /// on Microsoft.AspNetCore.* just to read one parameter.
    /// </summary>
    private static string? GetQueryValue(string query, string key)
    {
        if (string.IsNullOrEmpty(query)) return null;
        foreach (var pair in query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var idx = pair.IndexOf('=');
            var k = idx >= 0 ? pair[..idx] : pair;
            if (string.Equals(Uri.UnescapeDataString(k), key, StringComparison.OrdinalIgnoreCase))
            {
                if (idx < 0) return string.Empty;
                var rawValue = pair[(idx + 1)..];
                // Query-string values encode space as '+'. Convert '+' to space BEFORE percent-decoding,
                // otherwise the echoed token won't match what Graph sent.
                return Uri.UnescapeDataString(rawValue.Replace('+', ' '));
            }
        }
        return null;
    }
}