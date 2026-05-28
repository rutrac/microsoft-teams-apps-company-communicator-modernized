// <copyright file="TeamsConversationOrchestrator.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Utilities;

    /// <summary>
    /// Teams conversation orchestrator.
    /// </summary>
    public static class TeamsConversationOrchestrator
    {
        /// <summary>
        /// Run orchestrator.
        /// </summary>
        /// <param name="context">Durable orchestration context.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.TeamsConversationOrchestrator)]
        public static async Task RunOrchestrator(
            [OrchestrationTrigger] TaskOrchestrationContext context)
        {
            var log = context.CreateReplaySafeLogger(nameof(TeamsConversationOrchestrator));
            var batchPartitionKey = context.GetInput<string>();
            var notificationId = PartitionKeyUtility.GetNotificationIdFromBatchPartitionKey(batchPartitionKey);

            if (!context.IsReplaying)
            {
                log.LogInformation($"About to get pending recipients (with no conversation id in database).");
            }

            var recipients = await context.CallActivityAsync<IEnumerable<SentNotificationDataEntity>>(
                FunctionNames.GetPendingRecipientsActivity,
                batchPartitionKey,
                FunctionSettings.DefaultRetryOptions);

            var count = recipients.ToList().Count;
            if (count == 0)
            {
                log.LogInformation("No pending recipients.");
                return;
            }

            if (!context.IsReplaying)
            {
                log.LogInformation($"About to create 1:1 conversations with {count} recipients.");
            }

            var tasks = new List<Task>();
            foreach (var recipient in recipients)
            {
                recipient.PartitionKey = notificationId;

                var task = context.CallActivityAsync(
                    FunctionNames.TeamsConversationActivity,
                    (notificationId, batchPartitionKey, recipient),
                    FunctionSettings.DefaultRetryOptions);
                tasks.Add(task);
            }

            await Task.WhenAll(tasks);
        }
    }
}
