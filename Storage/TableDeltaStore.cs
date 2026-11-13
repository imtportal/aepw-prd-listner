using Azure;
using Azure.Data.Tables;

namespace Aepw.Evidence.Listener.Storage;

/// <summary>Azure Table Storage implementation, authenticated with Managed Identity (no keys).</summary>
public sealed class TableDeltaStore : IDeltaStore
{
    private const string TableName = "DeltaLinks";
    private const string Partition = "drive";
    private const string CTagPartition = "ctag";
    private readonly TableClient _table;

    public TableDeltaStore(TableServiceClient serviceClient)
    {
        _table = serviceClient.GetTableClient(TableName);
        _table.CreateIfNotExists();
    }

    public async Task<string?> GetDeltaLinkAsync(string driveId, CancellationToken ct = default)
    {
        try
        {
            var e = await _table.GetEntityAsync<TableEntity>(Partition, Key(driveId), cancellationToken: ct);
            return e.Value.GetString("DeltaLink");
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            return null;   // first run for this drive
        }
    }

    public async Task SaveDeltaLinkAsync(string driveId, string? deltaLink, CancellationToken ct = default)
    {
        var entity = new TableEntity(Partition, Key(driveId)) { ["DeltaLink"] = deltaLink };
        await _table.UpsertEntityAsync(entity, TableUpdateMode.Replace, ct);
    }

    public async Task<string?> GetCTagAsync(string driveId, string itemId, CancellationToken ct = default)
    {
        try
        {
            var e = await _table.GetEntityAsync<TableEntity>(CTagPartition, CTagKey(driveId, itemId), cancellationToken: ct);
            return e.Value.GetString("CTag");
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            return null;
        }
    }

    public async Task SaveCTagAsync(string driveId, string itemId, string cTag, CancellationToken ct = default)
    {
        var entity = new TableEntity(CTagPartition, CTagKey(driveId, itemId)) { ["CTag"] = cTag };
        await _table.UpsertEntityAsync(entity, TableUpdateMode.Replace, ct);
    }

    public async Task RemoveCTagAsync(string driveId, string itemId, CancellationToken ct = default)
    {
        try
        {
            await _table.DeleteEntityAsync(CTagPartition, CTagKey(driveId, itemId), cancellationToken: ct);
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            // already gone — no-op
        }
    }

    // RowKeys cannot contain '/', '\\', '#', '?' — encode all parts defensively.
    private static string Key(string driveId) => Uri.EscapeDataString(driveId);
    private static string CTagKey(string driveId, string itemId) =>
        $"{Uri.EscapeDataString(driveId)}|{Uri.EscapeDataString(itemId)}";
}
