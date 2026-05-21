// <copyright file="TeamDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData
{
    using global::Azure.Data.Tables;

    /// <summary>
    /// Teams data entity class.
    /// This entity holds the information about a team.
    /// </summary>
    public class TeamDataEntity : ITableEntity
    {
        public string PartitionKey { get; set; }
        public string RowKey { get; set; }
        public global::System.DateTimeOffset? Timestamp { get; set; }
        public global::Azure.ETag ETag { get; set; }
        /// <summary>
        /// Gets or sets the team id.
        /// </summary>
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets the name of the team.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the service url for the team.
        /// </summary>
        public string ServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets tenant id for the team.
        /// </summary>
        public string TenantId { get; set; }
    }
}
