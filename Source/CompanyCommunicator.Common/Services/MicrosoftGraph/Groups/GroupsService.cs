// <copyright file="GroupsService.cs" company="Microsoft">
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
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;

    /// <summary>
    /// Groups Service.
    /// </summary>
    internal class GroupsService : IGroupsService
    {
        private readonly GraphServiceClient graphServiceClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupsService"/> class.
        /// </summary>
        /// <param name="graphServiceClient">graph service client.</param>
        internal GroupsService(GraphServiceClient graphServiceClient)
        {
            this.graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(nameof(graphServiceClient));
        }

        private int MaxResultCount { get; set; } = 25;

        private int MaxRetry { get; set; } = 2;

        /// <summary>
        /// get groups by ids.
        /// </summary>
        /// <param name="groupIds">list of group ids.</param>
        /// <returns>list of groups.</returns>
        public async IAsyncEnumerable<Group> GetByIdsAsync(IEnumerable<string> groupIds)
        {
            foreach (var id in groupIds)
            {
                var group = await this.graphServiceClient
                                .Groups[id]
                                .GetAsync(req =>
                                {
                                    req.QueryParameters.Select = new[]
                                    {
                                        "id", "mail", "displayName", "visibility",
                                    };
                                });
                if (group != null)
                {
                    yield return group;
                }
            }
        }

        /// <summary>
        /// check if list has hidden membership group.
        /// </summary>
        /// <param name="groupIds">list of group ids.</param>
        /// <returns>boolean.</returns>
        public async Task<bool> ContainsHiddenMembershipAsync(IEnumerable<string> groupIds)
        {
            var groups = this.GetByIdsAsync(groupIds);
            await foreach (var group in groups)
            {
                if (group.IsHiddenMembership())
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Search M365 groups,distribution groups, security groups based on query.
        /// </summary>
        /// <param name="query">query param.</param>
        /// <returns>list of group.</returns>
        public async Task<IList<Group>> SearchAsync(string query)
        {
            query = Uri.EscapeDataString(query);
            var groupList = await this.SearchM365GroupsAsync(query, this.MaxResultCount);
            groupList.AddRange(await this.AddDistributionGroupAsync(query, this.MaxResultCount - groupList.Count()));
            groupList.AddRange(await this.AddSecurityGroupAsync(query, this.MaxResultCount - groupList.Count()));
            return groupList;
        }

        private async Task<List<Group>> SearchM365GroupsAsync(string query, int resultCount, bool includeHiddenMembership = false)
        {
            string filterQuery = $"groupTypes/any(c:c eq 'Unified') and mailEnabled eq true and (startsWith(mail,'{query}') or startsWith(displayName,'{query}'))";
            var groupsPaged = await this.SearchInternalAsync(filterQuery, resultCount);
            if (includeHiddenMembership)
            {
                return groupsPaged?.Value?.ToList() ?? new List<Group>();
            }

            var groupList = groupsPaged?.Value?
                .Where(group => !group.IsHiddenMembership())
                .ToList() ?? new List<Group>();

            while (groupsPaged?.OdataNextLink != null && groupList.Count() < resultCount)
            {
                groupsPaged = await this.graphServiceClient.Groups.WithUrl(groupsPaged.OdataNextLink).GetAsync();
                groupList.AddRange(groupsPaged?.Value?.Where(group => !group.IsHiddenMembership()) ?? Enumerable.Empty<Group>());
            }

            return groupList.Take(resultCount).ToList();
        }

        private async Task<IEnumerable<Group>> AddDistributionGroupAsync(string query, int resultCount)
        {
            if (resultCount == 0)
            {
                return new List<Group>();
            }

            string filterforDL = $"mailEnabled eq true and (startsWith(mail,'{query}') or startsWith(displayName,'{query}'))";
            var distributionGroups = await this.SearchInternalAsync(filterforDL, resultCount);

            var distributionGroupList = distributionGroups?.Value?
                .Where(dg => dg.GroupTypes.IsNullOrEmpty()).ToList() ?? new List<Group>();

            while (distributionGroups?.OdataNextLink != null && distributionGroupList.Count() < resultCount)
            {
                distributionGroups = await this.graphServiceClient.Groups.WithUrl(distributionGroups.OdataNextLink).GetAsync();
                distributionGroupList.AddRange(distributionGroups?.Value?.Where(dg => dg.GroupTypes.IsNullOrEmpty()) ?? Enumerable.Empty<Group>());
            }

            return distributionGroupList.Take(resultCount);
        }

        private async Task<IEnumerable<Group>> AddSecurityGroupAsync(string query, int resultCount)
        {
            if (resultCount == 0)
            {
                return new List<Group>();
            }

            string filterforSG = $"mailEnabled eq false and securityEnabled eq true and startsWith(displayName,'{query}')";
            var sgGroups = await this.SearchInternalAsync(filterforSG, resultCount);
            return sgGroups?.Value?.Take(resultCount) ?? Enumerable.Empty<Group>();
        }

        private async Task<Microsoft.Graph.Models.GroupCollectionResponse> SearchInternalAsync(string filterQuery, int resultCount)
        {
            return await this.graphServiceClient
                .Groups
                .GetAsync(req =>
                {
                    req.QueryParameters.Filter = filterQuery;
                    req.QueryParameters.Select = new[] { "id", "mail", "displayName", "visibility", "groupTypes" };
                    req.QueryParameters.Top = resultCount;
                });
        }
    }
}
