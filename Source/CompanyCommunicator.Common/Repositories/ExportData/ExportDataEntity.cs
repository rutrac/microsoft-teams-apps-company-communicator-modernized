// <copyright file="ExportDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ExportData
{
    using System;
    using global::Azure.Data.Tables;

    /// <summary>
    /// Export notification entity class.
    /// This entity holds all of the information about export.
    /// </summary>
    public class ExportDataEntity : ITableEntity
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
        /// Gets or sets the file name for the export data.
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// Gets or sets the response id of the File Consent Card.
        /// </summary>
        public string FileConsentId { get; set; }

        /// <summary>
        /// Gets or sets the DateTime of exporting the notification.
        /// </summary>
        public DateTime? SentDate { get; set; }

        /// <summary>
        /// Gets or sets the status of the export.
        /// </summary>
        public string Status { get; set; }
    }
}
