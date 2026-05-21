// <copyright file="GlobalSendingNotificationDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using System;
    using global::Azure.Data.Tables;

    /// <summary>
    /// Entity that holds metadata for all sending operations.
    /// This data is shared by all sending functions for all notifications.
    /// </summary>
    public class GlobalSendingNotificationDataEntity : ITableEntity
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
        /// Gets or sets a DateTime that sending of a notification can be retried/resumed.
        /// This is used to trigger a delay for all notifications if the bot is
        /// currently in a long term throttled state.
        /// After this given time, the sending function will attempt sending again to see
        /// if the bot is still in a throttled state.
        /// </summary>
        public DateTime? SendRetryDelayTime { get; set; }
    }
}
