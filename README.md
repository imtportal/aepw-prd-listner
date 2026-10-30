# Function App #1 — SharePoint Webhook Listener (standalone, ready to publish)

Self-contained .NET 8 isolated Function App. Hosts two functions:

| Function | Trigger | Role |
|---|---|---|
| `SharePointWebhook` | HTTP (POST) | Validation handshake, fast 202 ACK, Graph **delta** to find changed items, enqueue a **reference** per item to Service Bus |
| `SubscriptionRenewal` | Timer (daily 02:00 UTC) | Renew the SharePoint change subscription before it expires (~30-day cap) |

No project references, no secrets in code — all auth is Managed Identity. **This build has been compiled
(`dotnet build -c Release`) with 0 errors/0 warnings** against Microsoft.Graph 5.61.0.

## Files

```
Aepw.Evidence.Listener.csproj   packages + isolated-worker config
Program.cs                      DI wiring (Graph, Service Bus sender, Table delta store) via Managed Identity
host.json                       logging + App Insights sampling
local.settings.json             LOCAL ONLY template (not published; listed in .funcignore)
.funcignore / .gitignore
Models/ProcessingRequest.cs         queue message shape (reference, not bytes)
Models/GraphNotificationEnvelope.cs Graph notification envelope
Storage/IDeltaStore.cs              delta-link abstraction
Storage/TableDeltaStore.cs          Table Storage impl (Managed Identity)
Functions/SharePointWebhookListener.cs   Function #1
Functions/SubscriptionRenewal.cs         Function #3
```

## Prerequisites already in place (from the earlier infra/runbook)

- The Function App resource exists with a **system-assigned Managed Identity**, and that identity has:
  Graph **Sites.Selected** + per-site *write*, **Service Bus Data Sender**, **Storage Table Data Contributor**,
  **Storage Blob Data Owner**.
- Service Bus queue `evidence-queue` exists with duplicate detection enabled.
- Storage account exists (the app auto-creates the `DeltaLinks` table on first run).

## App settings to configure (in Azure, on this Function App)

| Setting | Example | Notes |
|---|---|---|
| `ServiceBusFullyQualifiedNamespace` | `sb-aepwevid.servicebus.windows.net` | MI auth, not a connection string |
| `ServiceBusQueueName` | `evidence-queue` | |
| `StorageAccountTableEndpoint` | `https://staepwevid.table.core.windows.net` | MI auth |
| `GraphSiteId` | `contoso.sharepoint.com,<guid>,<guid>` | from Graph |
| `GraphDriveId` | `b!...` | the Evidence library drive |
| `WebhookClientState` | `<random-guid>` | shared secret verifying inbound notifications |
| `SubscriptionId` | `<created-subscription-id>` | written back after webhook registration |
| `APPLICATIONINSIGHTS_CONNECTION_STRING` | `InstrumentationKey=...` | telemetry |

The runtime settings (`FUNCTIONS_WORKER_RUNTIME=dotnet-isolated`, `FUNCTIONS_EXTENSION_VERSION=~4`,
`AzureWebJobsStorage`) are set by the Bicep/infra deployment.

## Publish (directly into the Function App)

```bash
# from this folder
dotnet build -c Release

# Option A — Functions Core Tools (recommended)
func azure functionapp publish <listener-app-name> --dotnet-isolated

# Option B — zip deploy
dotnet publish -c Release -o ./publish
cd publish && zip -r ../app.zip . && cd ..
az functionapp deployment source config-zip \
  -g <resource-group> -n <listener-app-name> --src app.zip
```

Set the app settings (once):

```bash
CLIENT_STATE=$(uuidgen)
az functionapp config appsettings set -g <rg> -n <listener-app-name> --settings \
  ServiceBusFullyQualifiedNamespace="sb-aepwevid.servicebus.windows.net" \
  ServiceBusQueueName="evidence-queue" \
  StorageAccountTableEndpoint="https://staepwevid.table.core.windows.net" \
  GraphSiteId="<site-id>" GraphDriveId="<drive-id>" \
  WebhookClientState="$CLIENT_STATE"
```

## Register the webhook LAST (after this app is live)

Graph validates the callback URL at creation, so the app must already be deployed and reachable.

```bash
TOKEN=$(az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv)
curl -s -X POST https://graph.microsoft.com/v1.0/subscriptions \
  -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
  -d '{
        "changeType": "updated",
        "notificationUrl": "https://<listener-app-name>.azurewebsites.net/api/SharePointWebhook",
        "resource": "drives/<drive-id>/root",
        "expirationDateTime": "'$(date -u -d "+28 days" +%Y-%m-%dT%H:%M:%SZ)'",
        "clientState": "'$CLIENT_STATE'"
      }'
# copy the returned "id", then:
az functionapp config appsettings set -g <rg> -n <listener-app-name> \
  --settings SubscriptionId="<subscription-id>"
```

## Verify

- Handshake works: the POST above returns a subscription object with an `id` and future `expirationDateTime`.
- Upload a file to the SharePoint library → within seconds a message appears on `evidence-queue`
  (check the queue's active message count, or the Worker picking it up).
- `SubscriptionRenewal` shows a scheduled next run in the portal (Functions → SubscriptionRenewal).

## Local run (optional)

Fill `local.settings.json`, start Azurite (or point `AzureWebJobsStorage` at a real account), then:

```bash
func start
```

The HTTP endpoint will be `http://localhost:7071/api/SharePointWebhook`. Note the validation handshake
requires a public URL, so end-to-end webhook testing needs the deployed app (or a tunnel).
