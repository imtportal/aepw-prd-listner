using System.Security.Cryptography;
using System.Text;

namespace Aepw.Evidence.Listener.Models;

/// <summary>
/// Lightweight *reference* message placed on the Service Bus queue for the Worker to process.
/// We pass drive+item ids, never the file bytes, so the message stays well under Service Bus
/// size limits and the download happens in the Worker.
/// </summary>
public sealed class ProcessingRequest
{
    public string CorrelationId { get; set; } = Guid.NewGuid().ToString();
    public string EventType { get; set; } = "created";   // created | updated | deleted
    public string SiteId { get; set; } = string.Empty;
    public string DriveId { get; set; } = string.Empty;
    public string ItemId { get; set; } = string.Empty;
    public string? ETag { get; set; }
    public string FileName { get; set; } = string.Empty;
    public long SizeBytes { get; set; }
    public string? WebUrl { get; set; }
    public DateTime ReceivedUtc { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// Native Service Bus dedup key (set as the message MessageId). Hashed to a fixed 64-char hex
    /// string because the raw drive:item:etag concatenation exceeds Service Bus's 128-char
    /// MessageId limit. Identical drive+item+etag always yields the same hash, so dedup still holds.
    /// </summary>
    public string DedupKey
    {
        get
        {
            var raw = $"{DriveId}:{ItemId}:{ETag}";
            var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(raw));
            return Convert.ToHexString(bytes);   // always 64 chars
        }
    }
}