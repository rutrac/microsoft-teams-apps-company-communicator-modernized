// <copyright file="SendingNotificationDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using global::Azure.Data.Tables;

    /// <summary>
    /// Sending notification entity class.
    /// This entity holds the information about the content for a notification
    /// that is either currently being sent or was previously sent.
    /// </summary>
    public class SendingNotificationDataEntity : ITableEntity
    {
        public string PartitionKey { get; set; }
        public string RowKey { get; set; }
        public global::System.DateTimeOffset? Timestamp { get; set; }
        public global::Azure.ETag ETag { get; set; }
        /// <summary>
        /// Gets or sets the notification id.
        /// </summary>
        public string NotificationId { get; set; }

        /// <summary>
        /// Gets or sets the content of the notification in serialized JSON form.
        /// </summary>
        public string Content { get; set; }
    }
}
