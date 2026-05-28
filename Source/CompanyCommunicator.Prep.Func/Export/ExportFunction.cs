// <copyright file="ExportFunction.cs" company="Microsoft">
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
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ExportData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.ExportQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.Export.Model;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.Export.Orchestrator;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend;
    using Newtonsoft.Json;

    /// <summary>
    /// Azure Function App triggered by messages from a Service Bus queue.
    /// Exports notification as a zip file for the admin.
    /// </summary>
    public class ExportFunction
    {
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly IExportDataRepository exportDataRepository;
        private readonly IStringLocalizer<Strings> localizer;
        private readonly ILogger<ExportFunction> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExportFunction"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository.</param>
        /// <param name="exportDataRepository">Export data repository.</param>
        /// <param name="localizer">Localization service.</param>
        /// <param name="logger">Logger.</param>
        public ExportFunction(
            INotificationDataRepository notificationDataRepository,
            IExportDataRepository exportDataRepository,
            IStringLocalizer<Strings> localizer,
            ILogger<ExportFunction> logger)
        {
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.exportDataRepository = exportDataRepository ?? throw new ArgumentNullException(nameof(exportDataRepository));
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
            this.logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Service Bus triggered entry point that kicks off the export orchestration.
        /// </summary>
        /// <param name="message">The Service Bus received message.</param>
        /// <param name="starter">Durable orchestration client.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function("CompanyCommunicatorExportFunction")]
        public async Task Run(
            [ServiceBusTrigger(ExportQueue.QueueName, Connection = ExportQueue.ServiceBusConnectionConfigurationKey)]
            ServiceBusReceivedMessage message,
            [DurableClient] DurableTaskClient starter)
        {
            var myQueueItem = message.Body.ToString();
            var log = this.logger;

            var messageContent = JsonConvert.DeserializeObject<ExportMessageQueueContent>(myQueueItem);
            var notificationId = messageContent.NotificationId;
            var sentNotificationDataEntity = await this.notificationDataRepository.GetAsync(
                partitionKey: NotificationDataTableNames.SentNotificationsPartition,
                rowKey: notificationId);
            var exportDataEntity = await this.exportDataRepository.GetAsync(messageContent.UserId, notificationId);
            exportDataEntity.FileName = this.GetFileName();
            var requirement = new ExportDataRequirement(sentNotificationDataEntity, exportDataEntity, messageContent.UserId);
            if (!requirement.IsValid())
            {
                log.LogError("Export data requirement is not valid.");
                return;
            }

            string instanceId = await starter.ScheduleNewOrchestrationInstanceAsync(
                FunctionNames.ExportOrchestration,
                requirement);

            log.LogInformation($"Started orchestration with ID = '{instanceId}'.");
        }

        private string GetFileName()
        {
            var guid = Guid.NewGuid().ToString();
            var fileName = this.localizer.GetString("FileName_ExportData");
            return $"{fileName}_{guid}.zip";
        }
    }
}
