using System.Text.Json.Serialization;

namespace Aepw.Evidence.Listener.Models;

/// <summary>Envelope that Microsoft Graph POSTs to the webhook endpoint on a change.</summary>
public sealed class GraphNotificationEnvelope
{
    [JsonPropertyName("value")]
    public List<GraphNotification> Value { get; set; } = new();
}

public sealed class GraphNotification
{
    [JsonPropertyName("subscriptionId")] public string? SubscriptionId { get; set; }
    [JsonPropertyName("clientState")]    public string? ClientState { get; set; }
    [JsonPropertyName("resource")]       public string? Resource { get; set; }
    [JsonPropertyName("changeType")]     public string? ChangeType { get; set; }
    [JsonPropertyName("tenantId")]       public string? TenantId { get; set; }
}
