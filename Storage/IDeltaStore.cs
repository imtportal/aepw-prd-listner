namespace Aepw.Evidence.Listener.Storage;

/// <summary>
/// Persists the Microsoft Graph delta link per drive so each webhook fire resumes from the last
/// processed change instead of re-reading the whole library. Graph notifications only say
/// "the drive changed" — the delta link is how we discover exactly which items changed.
/// </summary>
public interface IDeltaStore
{
    Task<string?> GetDeltaLinkAsync(string driveId, CancellationToken ct = default);
    Task SaveDeltaLinkAsync(string driveId, string? deltaLink, CancellationToken ct = default);
}
