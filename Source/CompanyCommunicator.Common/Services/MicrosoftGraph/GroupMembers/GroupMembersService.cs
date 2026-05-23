// <copyright file="GroupMembersService.cs" company="Microsoft">
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

    /// <summary>
    /// Group Members Service.
    /// This gets the groups transitive members.
    /// </summary>
    internal class GroupMembersService : IGroupMembersService
    {
        private readonly GraphServiceClient graphServiceClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupMembersService"/> class.
        /// </summary>
        /// <param name="graphServiceClient">graph service client.</param>
        internal GroupMembersService(GraphServiceClient graphServiceClient)
        {
            this.graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(nameof(graphServiceClient));
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<User>> GetGroupMembersAsync(string groupId)
        {
            var response = await this.graphServiceClient
                                    .Groups[groupId]
                                    .TransitiveMembers
                                    .GraphUser
                                    .GetAsync(req =>
                                    {
                                        req.QueryParameters.Top = GraphConstants.MaxPageSize;
                                    });

            var users = response?.Value?.ToList() ?? new List<User>();
            while (response?.OdataNextLink != null)
            {
                response = await this.graphServiceClient
                    .Groups[groupId]
                    .TransitiveMembers
                    .GraphUser
                    .WithUrl(response.OdataNextLink)
                    .GetAsync();
                users.AddRange(response?.Value ?? new List<User>());
            }

            return users;
        }
    }
}
