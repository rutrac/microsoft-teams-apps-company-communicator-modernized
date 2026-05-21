// <copyright file="ChannelDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ChannelData
{
    using global::Azure.Data.Tables;

    /// <summary>
    /// Group Association data entity class.
    /// This entity holds the information about a group association with channels in teams.
    /// </summary>
    public class ChannelDataEntity : ITableEntity
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
        /// Gets or sets the channel id to where the group is associated.
        /// </summary>
        public string ChannelId { get; set; }

        /// <summary>
        /// Gets or sets the title to be used on cards for this channel.
        /// </summary>
        public string ChannelTitle { get; set; }

        /// <summary>
        /// Gets or sets the image to be used on cards for this channel.
        /// </summary>
        public string ChannelImage { get; set; }
    }
}