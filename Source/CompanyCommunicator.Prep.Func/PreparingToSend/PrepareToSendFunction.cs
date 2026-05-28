// <copyright file="PrepareToSendFunction.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func
{
    using System;
    using System.Threading.Tasks;
    using global::Azure.Messaging.ServiceBus;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask.Client;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.PrepareToSendQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend;
    using Newtonsoft.Json;

    /// <summary>
    /// Azure Function App triggered by messages from a Service Bus queue. <see cref="PrepareToSendQueue.QueueName"/>.
    /// </summary>
    public class PrepareToSendFunction
    {
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ILogger<PrepareToSendFunction> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="PrepareToSendFunction"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository.</param>
        /// <param name="logger">Logger.</param>
        public PrepareToSendFunction(
            INotificationDataRepository notificationDataRepository,
            ILogger<PrepareToSendFunction> logger)
        {
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Service Bus triggered entry point that kicks off the prepare-to-send orchestration.
        /// </summary>
        /// <param name="message">The Service Bus received message.</param>
        /// <param name="starter">Durable orchestration client.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.PrepareToSendFunction)]
        public async Task Run(
            [ServiceBusTrigger(PrepareToSendQueue.QueueName, Connection = PrepareToSendQueue.ServiceBusConnectionConfigurationKey)]
            ServiceBusReceivedMessage message,
            [DurableClient] DurableTaskClient starter)
        {
            var myQueueItem = message.Body.ToString();
            var log = this.logger;

            var queueMessageContent = JsonConvert.DeserializeObject<PrepareToSendQueueMessageContent>(myQueueItem);
            var notificationId = queueMessageContent.NotificationId;
            var sentNotificationDataEntity = await this.notificationDataRepository.GetAsync(
                partitionKey: NotificationDataTableNames.SentNotificationsPartition,
                rowKey: notificationId);

            if (sentNotificationDataEntity == null)
            {
                log.LogError($"Notification entity not found. Notification Id: {notificationId}");
                return;
            }

            string instanceId = await starter.ScheduleNewOrchestrationInstanceAsync(
                FunctionNames.PrepareToSendOrchestrator,
                sentNotificationDataEntity);

            log.LogInformation($"Started orchestration with ID = '{instanceId}'.");

            // Persist the instance id so the data function can later query orchestration status.
            sentNotificationDataEntity.FunctionInstancePayload = instanceId;
            await this.notificationDataRepository.InsertOrMergeAsync(sentNotificationDataEntity);
        }
    }
}
