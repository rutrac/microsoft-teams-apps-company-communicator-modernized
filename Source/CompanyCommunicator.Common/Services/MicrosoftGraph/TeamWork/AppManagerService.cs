// <copyright file="AppManagerService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Microsoft.Graph.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Policies;

    /// <summary>
    /// Manage Teams Apps for a user or a team.
    /// </summary>
    internal class AppManagerService : IAppManagerService
    {
        private readonly GraphServiceClient graphServiceClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="AppManagerService"/> class.
        /// </summary>
        /// <param name="graphServiceClient">V1 Graph service client.</param>
        internal AppManagerService(
            GraphServiceClient graphServiceClient)
        {
            this.graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(nameof(graphServiceClient));
        }

        /// <inheritdoc/>
        public async Task InstallAppForUserAsync(string appId, string userId)
        {
            if (string.IsNullOrWhiteSpace(appId))
            {
                throw new ArgumentNullException(nameof(appId));
            }

            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }

            var userScopeTeamsAppInstallation = new UserScopeTeamsAppInstallation
            {
                AdditionalData = new Dictionary<string, object>()
                {
                    { "teamsApp@odata.bind", $"{GraphConstants.V1BaseUrl}/appCatalogs/teamsApps/{appId}" },
                },
            };

            var retryPolicy = PollyPolicy.GetGraphRetryPolicy(GraphConstants.MaxRetry);
            await retryPolicy.ExecuteAsync(async ct =>
                await this.graphServiceClient.Users[userId]
                    .Teamwork
                    .InstalledApps
                    .PostAsync(userScopeTeamsAppInstallation, cancellationToken: ct));
        }

        /// <inheritdoc/>
        public async Task InstallAppForTeamAsync(string appId, string teamId)
        {
            if (string.IsNullOrWhiteSpace(appId))
            {
                throw new ArgumentNullException(nameof(appId));
            }

            if (string.IsNullOrWhiteSpace(teamId))
            {
                throw new ArgumentNullException(nameof(teamId));
            }

            var teamsAppInstallation = new TeamsAppInstallation()
            {
                AdditionalData = new Dictionary<string, object>()
                {
                    { "teamsApp@odata.bind", $"{GraphConstants.V1BaseUrl}/appCatalogs/teamsApps/{appId}" },
                },
            };

            var retryPolicy = PollyPolicy.GetGraphRetryPolicy(GraphConstants.MaxRetry);
            await retryPolicy.ExecuteAsync(async ct =>
                await this.graphServiceClient.Teams[teamId]
                    .InstalledApps
                    .PostAsync(teamsAppInstallation, cancellationToken: ct));
        }

        /// <inheritdoc/>
        public async Task<bool> IsAppInstalledForUserAsync(string appId, string userId)
        {
            if (string.IsNullOrWhiteSpace(appId))
            {
                throw new ArgumentNullException(nameof(appId));
            }

            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }

            var retryPolicy = PollyPolicy.GetGraphRetryPolicy(GraphConstants.MaxRetry);
            var pagedApps = await retryPolicy.ExecuteAsync(async ct =>
                await this.graphServiceClient.Users[userId]
                    .Teamwork
                    .InstalledApps
                    .GetAsync(
                        req =>
                        {
                            req.QueryParameters.Expand = new[] { "teamsApp" };
                            req.QueryParameters.Filter = $"teamsApp/id eq '{appId}'";
                        },
                        cancellationToken: ct));

            return pagedApps?.Value?.Any() ?? false;
        }

        /// <inheritdoc/>
        public async Task<bool> IsAppInstalledForTeamAsync(string appId, string teamId)
        {
            if (string.IsNullOrWhiteSpace(appId))
            {
                throw new ArgumentNullException(nameof(appId));
            }

            if (string.IsNullOrWhiteSpace(teamId))
            {
                throw new ArgumentNullException(nameof(teamId));
            }

            var retryPolicy = PollyPolicy.GetGraphRetryPolicy(GraphConstants.MaxRetry);
            var pagedApps = await retryPolicy.ExecuteAsync(async ct =>
                await this.graphServiceClient.Teams[teamId]
                    .InstalledApps
                    .GetAsync(
                        req =>
                        {
                            req.QueryParameters.Expand = new[] { "teamsApp" };
                            req.QueryParameters.Filter = $"teamsApp/id eq '{appId}'";
                        },
                        cancellationToken: ct));

            return pagedApps?.Value?.Any() ?? false;
        }

        /// <inheritdoc/>
        public async Task<string> GetAppInstallationIdForUserAsync(string appId, string userId)
        {
            if (string.IsNullOrWhiteSpace(appId))
            {
                throw new ArgumentNullException(nameof(appId));
            }

            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }

            var retryPolicy = PollyPolicy.GetGraphRetryPolicy(GraphConstants.MaxRetry);
            var collection = await retryPolicy.ExecuteAsync(async ct =>
                await this.graphServiceClient.Users[userId]
                    .Teamwork
                    .InstalledApps
                    .GetAsync(
                        req =>
                        {
                            req.QueryParameters.Expand = new[] { "teamsApp" };
                            req.QueryParameters.Filter = $"teamsApp/id eq '{appId}'";
                        },
                        cancellationToken: ct));

            return collection?.Value?.FirstOrDefault()?.Id;
        }

        /// <inheritdoc/>
        public async Task<string> GetAppInstallationIdForTeamAsync(string appId, string teamId)
        {
            if (string.IsNullOrWhiteSpace(appId))
            {
                throw new ArgumentNullException(nameof(appId));
            }

            if (string.IsNullOrWhiteSpace(teamId))
            {
                throw new ArgumentNullException(nameof(teamId));
            }

            var retryPolicy = PollyPolicy.GetGraphRetryPolicy(GraphConstants.MaxRetry);
            var collection = await retryPolicy.ExecuteAsync(async ct =>
                await this.graphServiceClient.Teams[teamId]
                    .InstalledApps
                    .GetAsync(
                        req =>
                        {
                            req.QueryParameters.Expand = new[] { "teamsApp" };
                            req.QueryParameters.Filter = $"teamsApp/id eq '{appId}'";
                        },
                        cancellationToken: ct));

            return collection?.Value?.FirstOrDefault()?.Id;
        }
    }
}
