namespace Aepw.Evidence.Listener.Storage;

/// <summary>
/// Persists the Microsoft Graph delta link per drive so each webhook fire resumes from the last
/// processed change instead of re-reading the whole library. Graph notifications only say
/// "the drive changed" — the delta link is how we discover exactly which items changed.
///
/// Also persists the last-enqueued cTag per item so the Listener can skip metadata-only
/// settle events (ETag bumps that do not change file content) without dropping real changes.
/// </summary>
public interface IDeltaStore
{
    Task<string?> GetDeltaLinkAsync(string driveId, CancellationToken ct = default);
    Task SaveDeltaLinkAsync(string driveId, string? deltaLink, CancellationToken ct = default);

    Task<string?> GetCTagAsync(string driveId, string itemId, CancellationToken ct = default);
    Task SaveCTagAsync(string driveId, string itemId, string cTag, CancellationToken ct = default);
    Task RemoveCTagAsync(string driveId, string itemId, CancellationToken ct = default);
}
