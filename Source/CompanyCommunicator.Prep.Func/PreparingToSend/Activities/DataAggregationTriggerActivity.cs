// <copyright file="DataAggregationTriggerActivity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.DataQueue;

    /// <summary>
    /// Data aggregation trigger activity.
    ///
    /// Does following:
    /// 1. Updates notification (total recipient count).
    /// 2. Sends message to data queue.
    /// </summary>
    public class DataAggregationTriggerActivity
    {
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly IDataQueue dataQueue;
        private readonly int messageDelayInSeconds;
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="DataAggregationTriggerActivity"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository.</param>
        /// <param name="dataQueue">The data queue.</param>
        /// <param name="options">The data queue message options.</param>
        public DataAggregationTriggerActivity(
            INotificationDataRepository notificationDataRepository,
            IDataQueue dataQueue,
            IOptions<DataQueueMessageOptions> options,
            ILogger<DataAggregationTriggerActivity> logger)
        {
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.dataQueue = dataQueue ?? throw new ArgumentNullException(nameof(dataQueue));
            this.messageDelayInSeconds = options?.Value?.MessageDelayInSeconds ?? throw new ArgumentNullException(nameof(options));
            this.logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Does following:
        /// 1. Updates notification (total recipient count).
        /// 2. Sends message to data queue.
        /// </summary>
        /// <param name="input">Input.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.DataAggregationTriggerActivity)]
        public async Task RunAsync(
            [ActivityTrigger](string notificationId, int recipientCount) input)
        {
            if (input.notificationId == null)
            {
                throw new ArgumentNullException(nameof(input.notificationId));
            }

            if (input.recipientCount <= 0)
            {
                throw new ArgumentOutOfRangeException($"Recipient count should be > 0. Value: {input.recipientCount}");
            }

            // Update notification.
            await this.UpdateNotification(input.notificationId, input.recipientCount);

            // Send message to data queue.
            var messageDelay = new TimeSpan(0, 0, this.messageDelayInSeconds);
            await this.dataQueue.SendMessageAsync(input.notificationId, messageDelay);
        }

        /// <summary>
        /// Update notification data (total recipient count).
        /// </summary>
        /// <param name="notificationId">Notification id.</param>
        /// <param name="recipientCount">Recipient count.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        private async Task UpdateNotification(string notificationId, int recipientCount)
        {
            var notificationDataEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.SentNotificationsPartition,
                notificationId);

            if (notificationDataEntity == null)
            {
                this.logger.LogError($"Notification entity not found. Notification Id: {notificationId}");
                return;
            }

            notificationDataEntity.TotalMessageCount = recipientCount;

            await this.notificationDataRepository.CreateOrUpdateAsync(notificationDataEntity);
        }
    }
}
