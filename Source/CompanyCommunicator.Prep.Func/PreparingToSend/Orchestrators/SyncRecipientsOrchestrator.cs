// <copyright file="SyncRecipientsOrchestrator.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Recipients;

    /// <summary>
    /// Syncs target set of recipients to Sent notification table.
    /// </summary>
    public static class SyncRecipientsOrchestrator
    {
        /// <summary>
        /// Fetch recipients and store them in Azure storage.
        /// </summary>
        /// <param name="context">Durable orchestration context.</param>
        /// <returns><see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.SyncRecipientsOrchestrator)]
        public static async Task<RecipientsInfo> RunOrchestrator(
            [OrchestrationTrigger] TaskOrchestrationContext context)
        {
            var log = context.CreateReplaySafeLogger(nameof(SyncRecipientsOrchestrator));
            var notification = context.GetInput<NotificationDataEntity>();

            await context.CallActivityAsync(
                FunctionNames.UpdateNotificationStatusActivity,
                (notification.Id, NotificationStatus.SyncingRecipients),
                FunctionSettings.DefaultRetryOptions);

            if (notification.AllUsers)
            {
                return await context.CallActivityAsync<RecipientsInfo>(
                    FunctionNames.SyncAllUsersActivity,
                    notification,
                    FunctionSettings.DefaultRetryOptions);
            }

            if (notification.Rosters.Any())
            {
                return await FanOutFanInActivityAsync(context, FunctionNames.SyncTeamMembersActivity, notification.Rosters, notification.Id);
            }

            if (notification.Groups.Any())
            {
                return await FanOutFanInActivityAsync(context, FunctionNames.SyncGroupMembersActivity, notification.Groups, notification.Id);
            }

            if (notification.Teams.Any())
            {
                return await context.CallActivityAsync<RecipientsInfo>(
                    FunctionNames.SyncTeamsActivity,
                    notification,
                    FunctionSettings.DefaultRetryOptions);
            }

            if (notification.CsvUsers.Length > 0)
            {
                log.LogInformation("Processing CSV Users.");
                return await context.CallActivityAsync<RecipientsInfo>(
                    FunctionNames.SyncCSVActivity,
                    notification,
                    FunctionSettings.DefaultRetryOptions);
            }

            var errorMessage = $"Invalid audience select for notification id: {notification.Id}";
            log.LogError(errorMessage);
            throw new ArgumentException(errorMessage);
        }

        private static async Task<RecipientsInfo> FanOutFanInActivityAsync(TaskOrchestrationContext context, string functionName, IEnumerable<string> entities, string notificationId)
        {
            var tasks = new List<Task>();
            int index = 1;
            foreach (var entityId in entities)
            {
                var task = context.CallActivityAsync(
                    functionName,
                    (notificationId, entityId, index),
                    FunctionSettings.DefaultRetryOptions);
                tasks.Add(task);
                index++;
            }

            await Task.WhenAll(tasks);

            return await context.CallActivityAsync<RecipientsInfo>(
                FunctionNames.BatchRecipientsActivity,
                notificationId,
                FunctionSettings.DefaultRetryOptions);
        }
    }
}
