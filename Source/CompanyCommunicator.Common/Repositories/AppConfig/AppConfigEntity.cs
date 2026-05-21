// <copyright file="AppConfigEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories
{
    using global::Azure.Data.Tables;

    /// <summary>
    /// App configuration entity.
    /// </summary>
    public class AppConfigEntity : ITableEntity
    {
        /// <summary>
        /// Gets or sets the entity's partition key.
        /// </summary>
        public string PartitionKey { get; set; }

        /// <summary>
        /// Gets or sets the entity's row key.
        /// </summary>
        public string RowKey { get; set; }

        /// <summary>
        /// Gets or sets the entity's timestamp.
        /// </summary>
        public global::System.DateTimeOffset? Timestamp { get; set; }

        /// <summary>
        /// Gets or sets the entity's ETag.
        /// </summary>
        public global::Azure.ETag ETag { get; set; }

        /// <summary>
        /// Gets or sets the entity value.
        /// </summary>
        public string Value { get; set; }
    }
}
