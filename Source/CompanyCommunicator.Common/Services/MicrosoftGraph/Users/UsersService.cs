// <copyright file="UsersService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Microsoft.Graph.Models;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Users service.
    /// </summary>
    internal class UsersService : IUsersService
    {
        private const string TeamsLicenseId = "57ff2da0-773e-42df-b2af-ffb7a2317929";

        private readonly GraphServiceClient graphServiceClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="UsersService"/> class.
        /// </summary>
        /// <param name="graphServiceClient">graph service client.</param>
        internal UsersService(GraphServiceClient graphServiceClient)
        {
            this.graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(nameof(graphServiceClient));
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<User>> GetBatchByUserIds(IEnumerable<IEnumerable<string>> userIdsByGroups)
        {
            if (userIdsByGroups == null)
            {
                throw new ArgumentNullException(nameof(userIdsByGroups));
            }

            var users = new List<User>();
            var batches = this.GetBatchRequests(userIdsByGroups);
            foreach (var batchRequestContent in batches)
            {
                var response = await this.graphServiceClient
                    .Batch
                    .PostAsync(batchRequestContent);

                var responses = await response.GetResponsesAsync();

                foreach (string key in responses.Keys)
                {
                    HttpResponseMessage httpResponse = default;
                    try
                    {
                        httpResponse = responses[key];
                        if (httpResponse == null)
                        {
                            throw new ArgumentNullException(nameof(httpResponse));
                        }

                        httpResponse.EnsureSuccessStatusCode();
                        var responseContent = await httpResponse.Content.ReadAsStringAsync();
                        JObject content = JObject.Parse(responseContent);
                        var userstemp = content["value"]
                            .Children()
                            .OfType<JObject>()
                            .Select(obj => obj.ToObject<User>());
                        if (userstemp == null)
                        {
                            continue;
                        }

                        users.AddRange(userstemp);
                    }
                    finally
                    {
                        httpResponse?.Dispose();
                    }
                }
            }

            return users;
        }

        /// <inheritdoc/>
        public async IAsyncEnumerable<IEnumerable<User>> GetUsersAsync(string filter = null)
        {
            var graphResult = await this.graphServiceClient
                .Users
                .GetAsync(req =>
                {
                    req.QueryParameters.Filter = filter;
                    req.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName" };
                });

            yield return graphResult.Value;

            while (graphResult.OdataNextLink != null)
            {
                graphResult = await this.graphServiceClient.Users.WithUrl(graphResult.OdataNextLink).GetAsync();
                yield return graphResult.Value;
            }
        }

        /// <inheritdoc/>
        public async Task<User> GetUserAsync(string userId)
        {
            return await this.graphServiceClient
                .Users[userId]
                .GetAsync(req =>
                {
                    req.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName", "userType" };
                });
        }

        /// <inheritdoc/>
        public async Task<(IEnumerable<User>, string)> GetAllUsersAsync(string deltaLink = null)
        {
            var users = new List<User>();
            Microsoft.Graph.Users.Delta.DeltaGetResponse deltaResponse;

            if (string.IsNullOrEmpty(deltaLink))
            {
                deltaResponse = await this.graphServiceClient
                    .Users
                    .Delta
                    .GetAsync(req =>
                    {
                        req.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName", "userType" };
                        req.QueryParameters.Top = GraphConstants.MaxPageSize;
                    });
            }
            else
            {
                deltaResponse = await this.graphServiceClient.Users.Delta.WithUrl(deltaLink).GetAsync();
            }

            users.AddRange(deltaResponse.Value);

            while (deltaResponse.OdataNextLink != null)
            {
                deltaResponse = await this.graphServiceClient.Users.Delta.WithUrl(deltaResponse.OdataNextLink).GetAsync();
                users.AddRange(deltaResponse.Value);
            }

            return (users, deltaResponse.OdataDeltaLink);
        }

        /// <inheritdoc/>
        public async Task<bool> HasTeamsLicenseAsync(string userId)
        {
            if (string.IsNullOrEmpty(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }

            var licenseCollection = await this.graphServiceClient
                .Users[userId]
                .LicenseDetails
                .GetAsync(req =>
                {
                    req.QueryParameters.Top = GraphConstants.MaxPageSize;
                });

            if (this.HasTeamsLicense(licenseCollection.Value))
            {
                return true;
            }

            while (licenseCollection.OdataNextLink != null)
            {
                licenseCollection = await this.graphServiceClient
                    .Users[userId]
                    .LicenseDetails
                    .WithUrl(licenseCollection.OdataNextLink)
                    .GetAsync();

                if (this.HasTeamsLicense(licenseCollection.Value))
                {
                    return true;
                }
            }

            return false;
        }

        private string GetUserIdFilter(IEnumerable<string> userIds)
        {
            StringBuilder filterUserIds = new StringBuilder();
            foreach (var id in userIds)
            {
                if (!string.IsNullOrEmpty(filterUserIds.ToString()))
                {
                    filterUserIds.Append(" or ");
                }

                filterUserIds.Append($"id eq '{id}'");
            }

            return filterUserIds.ToString();
        }

        private IEnumerable<BatchRequestContent> GetBatchRequests(IEnumerable<IEnumerable<string>> userIdsByGroups)
        {
            var batches = new List<BatchRequestContent>();
            int maxNoBatchItems = 20;

            var batchRequestContent = new BatchRequestContent(this.graphServiceClient);
            int requestId = 1;

            foreach (var userIds in userIdsByGroups)
            {
                if (userIds.Count() == 0)
                {
                    continue;
                }

                if (userIds.Count() > 15)
                {
                    throw new InvalidOperationException("The id count should be less than or equal to 15");
                }

                var filterUserIds = this.GetUserIdFilter(userIds);
                var requestInfo = this.graphServiceClient.Users.ToGetRequestInformation(req =>
                {
                    req.QueryParameters.Filter = filterUserIds;
                    req.QueryParameters.Select = new[] { "id", "displayName", "userPrincipalName", "userType" };
                });
                batchRequestContent.AddBatchRequestStepAsync(requestInfo, requestId.ToString()).GetAwaiter().GetResult();

                if (batchRequestContent.BatchRequestSteps.Count() % maxNoBatchItems == 0)
                {
                    batches.Add(batchRequestContent);
                    batchRequestContent = new BatchRequestContent(this.graphServiceClient);
                }

                requestId++;
            }

            if (batchRequestContent.BatchRequestSteps.Count > 0 && batchRequestContent.BatchRequestSteps.Count < maxNoBatchItems)
            {
                batches.Add(batchRequestContent);
            }

            return batches;
        }

        private bool HasTeamsLicense(IEnumerable<Microsoft.Graph.Models.LicenseDetails> licenses)
        {
            foreach (var license in licenses)
            {
                if (license.ServicePlans == null)
                {
                    continue;
                }

                if (license.ServicePlans.Any(sp => string.Equals(sp.ServicePlanId?.ToString(), TeamsLicenseId)))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
