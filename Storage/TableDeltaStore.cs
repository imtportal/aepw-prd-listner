using Azure;
using Azure.Data.Tables;

namespace Aepw.Evidence.Listener.Storage;

/// <summary>Azure Table Storage implementation, authenticated with Managed Identity (no keys).</summary>
public sealed class TableDeltaStore : IDeltaStore
{
    private const string TableName = "DeltaLinks";
    private const string Partition = "drive";
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

    // RowKeys cannot contain '/', '\\', '#', '?' — encode the drive id defensively.
    private static string Key(string driveId) => Uri.EscapeDataString(driveId);
}
