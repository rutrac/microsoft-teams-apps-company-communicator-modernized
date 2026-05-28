// <copyright file="PrepareToSendOrchestrator.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Recipients;

    /// <summary>
    /// Prepare to Send orchestrator.
    /// </summary>
    public static class PrepareToSendOrchestrator
    {
        /// <summary>
        /// Kicks off the preparing to send process.
        /// </summary>
        /// <param name="context">Durable orchestration context.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        [Function(FunctionNames.PrepareToSendOrchestrator)]
        public static async Task RunOrchestrator(
            [OrchestrationTrigger] TaskOrchestrationContext context)
        {
            var log = context.CreateReplaySafeLogger(nameof(PrepareToSendOrchestrator));
            var notificationDataEntity = context.GetInput<NotificationDataEntity>();

            if (!context.IsReplaying)
            {
                log.LogInformation($"Start to prepare to send the notification {notificationDataEntity.Id}!");
            }

            try
            {
                if (!context.IsReplaying)
                {
                    log.LogInformation("About to store message content.");
                }

                await context.CallActivityAsync(
                    FunctionNames.StoreMessageActivity,
                    notificationDataEntity,
                    FunctionSettings.DefaultRetryOptions);

                if (!context.IsReplaying)
                {
                    log.LogInformation("About to sync recipients.");
                }

                var recipientsInfo = await context.CallSubOrchestratorAsync<RecipientsInfo>(
                    FunctionNames.SyncRecipientsOrchestrator,
                    notificationDataEntity,
                    FunctionSettings.DefaultRetryOptions);

                if (recipientsInfo.HasRecipientsPendingInstallation)
                {
                    if (!context.IsReplaying)
                    {
                        log.LogInformation("About to create 1:1 conversations for recipients if required.");
                    }

                    await context.CallActivityAsync(
                        FunctionNames.UpdateNotificationStatusActivity,
                        (recipientsInfo.NotificationId, NotificationStatus.InstallingApp),
                        FunctionSettings.DefaultRetryOptions);

                    await FanOutFanInSubOrchestratorAsync(context, FunctionNames.TeamsConversationOrchestrator, recipientsInfo);
                }

                if (!context.IsReplaying)
                {
                    log.LogInformation("About to send messages to send queue.");
                }

                await context.CallActivityAsync(
                    FunctionNames.UpdateNotificationStatusActivity,
                    (notificationDataEntity.Id, NotificationStatus.Sending),
                    FunctionSettings.DefaultRetryOptions);

                await context.CallActivityAsync(
                    FunctionNames.DataAggregationTriggerActivity,
                    (notificationDataEntity.Id, recipientsInfo.TotalRecipientCount),
                    FunctionSettings.DefaultRetryOptions);

                await FanOutFanInSubOrchestratorAsync(context, FunctionNames.SendQueueOrchestrator, recipientsInfo);

                log.LogInformation($"PrepareToSendOrchestrator successfully completed for notification: {notificationDataEntity.Id}!");
            }
            catch (Exception ex)
            {
                var errorMessage = $"PrepareToSendOrchestrator failed for notification: {notificationDataEntity.Id}. Exception Message: {ex.Message}";
                log.LogError(ex, errorMessage);

                await context.CallActivityAsync(
                    FunctionNames.HandleFailureActivity,
                    (notificationDataEntity, ex),
                    FunctionSettings.DefaultRetryOptions);
            }
        }

        private static async Task FanOutFanInSubOrchestratorAsync(TaskOrchestrationContext context, string functionName, RecipientsInfo recipientsInfo)
        {
            var tasks = new List<Task>();
            foreach (var batchKey in recipientsInfo.BatchKeys)
            {
                var task = context.CallSubOrchestratorAsync(
                    functionName,
                    batchKey,
                    FunctionSettings.DefaultRetryOptions);
                tasks.Add(task);
            }

            await Task.WhenAll(tasks);
        }
    }
}
