// <copyright file="AggregateSentNotificationDataService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Data.Func.Services.NotificationDataServices
{
    using System;
    using System.Threading.Tasks;
    using Azure.Data.Tables;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;

    /// <summary>
    /// A service that fetches and aggregates the Sent Notification Data results.
    /// </summary>
    public class AggregateSentNotificationDataService
    {
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="AggregateSentNotificationDataService"/> class.
        /// </summary>
        /// <param name="sentNotificationDataRepository">The sent notification data repository.</param>
        public AggregateSentNotificationDataService(ISentNotificationDataRepository sentNotificationDataRepository)
        {
            this.sentNotificationDataRepository = sentNotificationDataRepository;
        }

        /// <summary>
        /// Fetches all of the current known results for the Sent Notification and calculates the various totals
        /// as results.
        /// </summary>
        /// <param name="notificationId">The notification ID.</param>
        /// <param name="log">The logger.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<AggregatedSentNotificationDataResults> AggregateSentNotificationDataResultsAsync(
            string notificationId,
            ILogger log)
        {
            // PartitionKey eq notificationId AND DeliveryStatus ne 'null'
            // The SentNotificationDataEntity.DeliveryStatus property's default value is null.
            // After finished processing a recipient, the send function sets the property to one of the following
            // values, which indicates the delivery status: Succeeded, Failed, RecipientNotFound, Throttled, etc.
            var filter = $"(PartitionKey eq '{notificationId}') and (DeliveryStatus ne 'null')";

            try
            {
                var aggregatedResults = new AggregatedSentNotificationDataResults();

                await foreach (var sentNotification in this.sentNotificationDataRepository.Table
                    .QueryAsync<SentNotificationDataEntity>(filter))
                {
                    aggregatedResults.UpdateAggregatedResults(sentNotification);
                }

                return aggregatedResults;
            }
            catch (Exception e)
            {
                var errorMessage = $"{e.GetType()}: {e.Message}";
                log.LogError(e, $"ERROR: {errorMessage}");
                throw;
            }
        }
    }
}
