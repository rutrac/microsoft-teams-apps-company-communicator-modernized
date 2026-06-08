## Release Notes

Cumulative improvements in Company Communicator App.

### Version history

#### Modernized fork (v5.14 – v6.0)

The modernized fork is **fresh-deploy only** — there is no supported upgrade path from v5.13 or earlier.

| Release | Stable tag | Highlights |
|---|---|---|
| v6.0 | `v6.0.0-final` | Final modernized release. Wiki refreshed end-to-end for the modernized stack (.NET 8 isolated, Vite, Northstar, AFD, VNet+PEs, Premium V3); legacy upgrade guides and the upstream Scale/Extending pages removed. `it-IT` added to the default `supportedCultures` so fresh deploys ship the locale that v5.26.603 introduced. End-to-end smoke validated on cc-test34: draft+send, send to all users (immediate), and Graph admin-consent path documented. |
| v5.29 | `v5.29.0-legacy-removed` | Phase 6: removed all legacy migration code paths (`isUpgrade` flow, `AadOptionalClaims_Reset.json`, legacy `HttpManagementPayload` branches in `SentNotificationsController` and `UpdateNotificationDataService`, the five v2-v5 migration wiki guides). Fresh deploy is the only supported path. |
| v5.28 | `v5.28.0-stable` | Dependabot config (npm + nuget + github-actions); Playwright pure-UI smoke tests; TypeScript 4.9 → 5.9; moment 2.30 → dayjs 1.11 with on-demand locale loading. |
| v5.27.5 | `v5.27.5-defender-hardened` | Broader Azure CLI detection in `deploy.ps1` (winget, scoop, choco, user-scope, x64 MSI); pinned vulnerable NuGet transitives; parameterized Key Vault purge protection (off in sandbox, on in production). |
| v5.27.4 | `v5.27.4-northstar-restored` | Reverted the Fluent v9 client port; restored Northstar UI from `v5.27.2-hooks-stable`; re-checked-in the pre-built `ClientApp/build` artifact so Kudu source-control sync does not have to run npm. |
| v5.27.3 | `v5.27.3-fluent9-stable` | (Reverted in v5.27.4.) Phase 4 Fluent v9 migration — dropped Northstar, ported all message-authoring components to `@fluentui/react-components`. |
| v5.27.2 | `v5.27.2-hooks-stable` | Phase 3b: converted all class components to functional + hooks. React Router v5 → v6. `deploy.ps1` enforces `alwaysOn` + 64-bit worker on all sites to eliminate dotnet-isolated cold starts. |
| v5.27.1 | `v5.27.1-r18-stable` | React 18 migration. `.npmrc` adds `legacy-peer-deps=true` for Kudu's npm 10.x. |
| v5.27.0 | `v5.27.0-stable` | Phase 1: CRA → Vite client build. Friendly UX on `/signin-end` when redirected from the admin-consent flow. |
| v5.26.603 | `v5.26.603-stable` | Tenant-policy hardening for `deployIfNotExists` diagnostic-setting policies. Added `it-IT` (Italian) locale and `pt-PT` alias mirroring `pt-BR`. |
| v5.26.602 | `v5.26.602-stable` | Admin-consent prompt 30s timeout (default N, non-blocking); npm audit bumps (axios 0.28→1.16, i18next-http-backend 1.4→4.0, markdown-it 13→14); removed bogus `System.Net.Http 4.3.4` / `System.Text.RegularExpressions 4.3.1` direct refs; StyleCop `.editorconfig` with pragmatic suppressions; added private endpoints + DNS zones for Storage **table** and **queue** services (in addition to blob). |
| v5.26.6 (intermediate, untagged) | — | Unified retry helper for Graph/AAD calls + transcript log. Pinned transitive CVE deps. Removed unused `gitRepoUrl` / `gitBranch` params and dead ServiceBus / AppInsights `secret-resourceId` vars from ARM. Bot web app: `defaultAction=Allow` on main site (Teams traffic was being blocked); SCM stays `Deny`. Auto-accept `MsTeamsChannel` ToS for user + author bots after ARM (later reverted — not the blocker). |
| v5.26.5 (intermediate) | — | Adaptive Card: render red **IMPORTANT** banner when `IsImportant=true` (both client and server-built cards). FAQ entry explaining auto-created `NetworkWatcherRG`. Rewrote four PowerShell `if`-expression assignments to statement form to avoid a parser crash after a native `az` call. |
| v5.26.4 (intermediate) | — | Scrubbed sandbox identifiers from the public `parameters.json` template. Added VNet + Private Endpoints (Storage blob, Key Vault); IP allowlists + access restrictions; fixed `isSharedPlan` regression. Locked App Service Plan to **PremiumV3** (`P0v3` or `P1v3`); removed the sizing wizard. Removed certificate-auth + GCC deployment paths (out of scope for the modernized fork). |
| v5.14 – v5.26.3 (intermediate, untagged) | — | Graph v6 OData filter syntax fix (spaces, not `+`) in groups search. Serialize `AdaptiveCard` to `JObject` to avoid the SDK serializer stack overflow. `SendActivityAsync` instrumented with a probe + 20s timeout. Initial .NET 8 isolated functions move and Azure Front Door Standard adoption. |

#### Upstream version history (pre-fork)

|Release |Published to <br/> Microsoft Store |
|---|---|
| 5.1 | April 28, 2022
| 5.0 | Nov 10, 2021
| 4.1.5 | Sep 29, 2021
| 4.1.4 | Sep 14, 2021
| 4.1.3 | Jul 2, 2021
| 4.1.2 | Jun 25, 2021
| 4.1.1 | Jun 12, 2021
| 4.1 | Mar 19, 2021
| 4.1 | Mar 19, 2021
| 4.0 | Dec 30, 2020
| 3.0 | Oct 29, 2020
| 2.1 | Oct 16, 2020
| 2.0 | Aug 20, 2020
| 1.1 | Jun 08, 2020
| 1.0 | Dec 20, 2019

### Company Communicator feature release notes

#### 5.1 (April 28, 2022)
##### Changes introduced
- Ability to cancel a notification.
- Export installation errors.
- Arm fixes.

#### 5.0 (Nov 10, 2021)
##### Changes introduced
- Added Key Vault and Managed Identity.
- Support certificate authentication.
- Bug fix to resolve expired delta url. 

#### 4.1.5 (Sep 29, 2021)
##### Changes introduced
- Limit the size of the error and warning messages stored to 1024 characters.

#### 4.1.4 (Sep 14, 2021)
##### Changes introduced
- Support large number of users.
- Reduce memory usage.

#### 4.1.3 (Jul 2, 2021)
##### Changes introduced
- Export report for users who have left tenant.

#### 4.1.2 (Jun 25, 2021)
##### Changes introduced
- Exclude existing guest users with user app installed from receiving message.
- Identify UserType using export report functionality.
- Bug fix preventing proactive installations.

#### 4.1.1 (Jun 12, 2021)
##### Changes introduced
- Exclude guest users when sending message to:
  - Members of one or more Teams.
  - Members of one or more Groups.
- Bug fix with the author app interface in dark and high-contrast themes.
- Resolved potential out of memory errors when sending message to large audience.

#### 4.1 (Mar 19, 2021)
##### Changes introduced
- Locale support for multiple languages.
- Migration to fluent ui northstar.
- Migrating graph beta apis to v1.0.
- Improved Test coverage.

#### 4.0 (Dec 30, 2020)
##### Changes intoduced
- Separate Bots for User and Author.
- Automated deployment using Powershell script.
- Improved Test coverage.

#### 3.0 (Oct 29, 2020)
###### Changes introduced
- Proactive User app installation.
- Send message to all the users in a tenant.
- Multi-Locale support in backend and client application.
- Granular status updates after sending a message.
- Performance improvements.
- Quality and reliability fixes.

#### 2.1 (Oct 16, 2020)
###### Changes introduced
- Bug fix.
- Performance improvements.

#### 2.0 (Aug 20, 2020)
###### Changes introduced
- Send a message to an M365 group, SG or DG.
- Search an M365 group, SG or DG.
- Export data for messages sent.
- Updated to MSBuild v16.

#### 1.1 (Jun 08, 2020)
###### Changes introduced
- Upgraded NPM Packages.

#### 1.0 (Dec 20, 2019)
###### Changes introduced
- Company Communicator template released.
- Ability to draft/send messages.
- Ability to send a message to:
  - Members of one or more Teams.
  - General channel of one or more Teams.
  - All the users who install the User app.