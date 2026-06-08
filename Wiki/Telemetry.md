# Telemetry

The Company Communicator app logs telemetry to [Azure Application Insights](https://learn.microsoft.com/azure/azure-monitor/app/app-insights-overview). Open the Application Insights resource provisioned by the template (named after your `baseResourceName`) to see live metrics, request traces, dependency calls, and exceptions across all four sites (web app + 3 function apps).

## How telemetry is wired

- **App Insights connection string** — stored in Key Vault and surfaced to each site as `APPLICATIONINSIGHTS_CONNECTION_STRING` via a Key Vault reference. The legacy `APPINSIGHTS_INSTRUMENTATIONKEY` setting is no longer used.
- **ASP.NET Core 8 web app** — uses `Microsoft.ApplicationInsights.AspNetCore` for incoming HTTP requests, dependency tracking, and exception capture.
- **.NET 8 isolated Functions (v4)** — use the Functions host's built-in Application Insights integration. Custom telemetry uses `ILogger<T>` injected via DI; the host pipes log messages to App Insights automatically.
- **Bot activity** — the Bot Framework adapter integrates with Application Insights as described in [Bot Framework analytics behind the scenes](https://blog.botframework.com/2019/03/21/bot-analytics-behind-the-scenes/).
- **Diagnostic settings** — `deploy.ps1` can pre-create tenant-managed `deployIfNotExists` diagnostic settings (see the `policyDiagnosticSettings` parameter in `parameters.json`) so that platform logs flow to your Log Analytics workspace from the first ARM deploy onwards, without triggering a remediation race.

## Event categories

| Category | Source | What you'll see |
|---|---|---|
| `Trace` | `ILogger<T>.LogInformation` / `LogDebug` | Application events — prep orchestration progress, send loop iterations, recipient sync milestones. |
| `Request` | ASP.NET Core / Functions host | HTTP requests to the web app and HTTP-triggered functions. |
| `Dependency` | App Insights auto-collection | Outbound calls to Microsoft Graph, Azure Storage tables/queues/blobs, Service Bus, Key Vault. Useful for chasing throttling and timeouts. |
| `Exception` | `ILogger<T>.LogError(ex, …)` and unhandled exceptions | Stack traces from failed sends, Graph 4xx/5xx, Service Bus dead-letters, etc. |
| `Custom Event` | `TelemetryClient.TrackEvent` (selected hot paths only) | High-level lifecycle: notification drafted, send started, send completed. |

## Useful Kusto queries

Drop these into the App Insights **Logs** blade:

```kusto
// Last 24h of failed sends per function app
exceptions
| where timestamp > ago(24h)
| where cloud_RoleName endswith "-function" or cloud_RoleName endswith "-prep-function" or cloud_RoleName endswith "-data-function"
| summarize count() by cloud_RoleName, type, outerMessage
| order by count_ desc
```

```kusto
// Service Bus dependency failures
dependencies
| where timestamp > ago(24h)
| where type == "Azure Service Bus" and success == false
| project timestamp, cloud_RoleName, name, resultCode, duration, operation_Id
| order by timestamp desc
```

```kusto
// End-to-end trace for a single notification
let notif = "<notification-id>";
union traces, exceptions, dependencies
| where customDimensions.NotificationId == notif
| order by timestamp asc
```