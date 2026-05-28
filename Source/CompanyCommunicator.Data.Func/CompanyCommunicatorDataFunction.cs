// <copyright file="CompanyCommunicatorDataFunction.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Data.Func
{
    using System;
    using System.Threading.Tasks;
    using global::Azure.Messaging.ServiceBus;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask.Client;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.DataQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Data.Func.Services.NotificationDataServices;
    using Newtonsoft.Json;

    /// <summary>
    /// Azure Function App triggered by messages from a Service Bus queue
    /// Used for incrementing results for a sent notification.
    /// </summary>
    public class CompanyCommunicatorDataFunction
    {
        private static readonly double TenMinutes = 10;

        private readonly INotificationDataRepository notificationDataRepository;
        private readonly AggregateSentNotificationDataService aggregateSentNotificationDataService;
        private readonly UpdateNotificationDataService updateNotificationDataService;
        private readonly IDataQueue dataQueue;
        private readonly double firstTenMinutesRequeueMessageDelayInSeconds;
        private readonly double requeueMessageDelayInSeconds;
        private readonly ILogger<CompanyCommunicatorDataFunction> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyCommunicatorDataFunction"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">The notification data repository.</param>
        /// <param name="aggregateSentNotificationDataService">The service to aggregate the Sent
        /// Notification Data results.</param>
        /// <param name="updateNotificationDataService">The service to update the notification totals.</param>
        /// <param name="dataQueue">The data queue.</param>
        /// <param name="dataQueueMessageOptions">The data queue message options.</param>
        /// <param name="logger">Logger.</param>
        public CompanyCommunicatorDataFunction(
            INotificationDataRepository notificationDataRepository,
            AggregateSentNotificationDataService aggregateSentNotificationDataService,
            UpdateNotificationDataService updateNotificationDataService,
            IDataQueue dataQueue,
            IOptions<DataQueueMessageOptions> dataQueueMessageOptions,
            ILogger<CompanyCommunicatorDataFunction> logger)
        {
            this.notificationDataRepository = notificationDataRepository;
            this.aggregateSentNotificationDataService = aggregateSentNotificationDataService;
            this.updateNotificationDataService = updateNotificationDataService;
            this.dataQueue = dataQueue;
            this.firstTenMinutesRequeueMessageDelayInSeconds =
                dataQueueMessageOptions.Value.FirstTenMinutesRequeueMessageDelayInSeconds;
            this.requeueMessageDelayInSeconds =
                dataQueueMessageOptions.Value.RequeueMessageDelayInSeconds;
            this.logger = logger;
        }

        /// <summary>
        /// Azure Function App triggered by messages from a Service Bus queue
        /// Used for aggregating results for a sent notification.
        /// </summary>
        /// <param name="message">The Service Bus received message.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function("CompanyCommunicatorDataFunction")]
        public async Task Run(
            [ServiceBusTrigger(
                DataQueue.QueueName,
                Connection = DataQueue.ServiceBusConnectionConfigurationKey)]
            ServiceBusReceivedMessage message,
            [DurableClient] DurableTaskClient durableTaskClient)
        {
            var myQueueItem = message.Body.ToString();
            var log = this.logger;
            var messageContent = JsonConvert.DeserializeObject<DataQueueMessageContent>(myQueueItem);

            var notificationDataEntity = await this.notificationDataRepository.GetAsync(
                partitionKey: NotificationDataTableNames.SentNotificationsPartition,
                rowKey: messageContent.NotificationId);

            // If notification is already marked complete, then there is nothing left to do for the data queue trigger.
            if (!notificationDataEntity.IsCompleted())
            {
                string orchestrationStatus = string.Empty;
                if (notificationDataEntity.Status.Equals(NotificationStatus.Canceling.ToString()))
                {
                    orchestrationStatus = await this.updateNotificationDataService.GetOrchestrationStatusAsync(notificationDataEntity.FunctionInstancePayload, durableTaskClient);
                }

                // Get all of the result counts (Successes, Failures, etc.) from the Sent Notification Data.
                var aggregatedSentNotificationDataResults = await this.aggregateSentNotificationDataService
                .AggregateSentNotificationDataResultsAsync(messageContent.NotificationId, log);

                // Use these counts to update the Notification Data accordingly.
                var notificationDataEntityUpdate = await this.updateNotificationDataService
                    .UpdateNotificationDataAsync(
                        notificationId: messageContent.NotificationId,
                        orchestrationStatus: orchestrationStatus,
                        shouldForceCompleteNotification: messageContent.ForceMessageComplete,
                        totalExpectedNotificationCount: notificationDataEntity.TotalMessageCount,
                        aggregatedSentNotificationDataResults: aggregatedSentNotificationDataResults,
                        log: log);

                // If the notification is still not in a completed state, then requeue the Data Queue trigger
                // message with a delay in order to aggregate the results again.
                if (!notificationDataEntityUpdate.IsCompleted())
                {
                    // Requeue data aggregation trigger message with a delay to calculate the totals again.
                    var dataQueueTriggerMessage = new DataQueueMessageContent
                    {
                        NotificationId = messageContent.NotificationId,
                        ForceMessageComplete = false,
                    };

                    var dataQueueTriggerMessageDelayInSeconds =
                        DateTime.UtcNow <= notificationDataEntity.SendingStartedDate + TimeSpan.FromMinutes(CompanyCommunicatorDataFunction.TenMinutes)
                            ? this.firstTenMinutesRequeueMessageDelayInSeconds
                            : this.requeueMessageDelayInSeconds;

                    await this.dataQueue.SendDelayedAsync(dataQueueTriggerMessage, dataQueueTriggerMessageDelayInSeconds);
                }
            }
        }
    }
}
