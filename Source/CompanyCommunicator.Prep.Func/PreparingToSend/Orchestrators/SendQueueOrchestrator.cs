// <copyright file="SendQueueOrchestrator.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend.Orchestrators
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Utilities;

    /// <summary>
    /// Send Queue orchestrator.
    /// </summary>
    public static class SendQueueOrchestrator
    {
        /// <summary>
        /// SendQueueSubOrchestrator function.
        /// </summary>
        /// <param name="context">Durable orchestration context.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.SendQueueOrchestrator)]
        public static async Task RunOrchestrator(
            [OrchestrationTrigger] TaskOrchestrationContext context)
        {
            var log = context.CreateReplaySafeLogger(nameof(SendQueueOrchestrator));
            var batchPartitionKey = context.GetInput<string>();
            var notificationId = PartitionKeyUtility.GetNotificationIdFromBatchPartitionKey(batchPartitionKey);
            var batchId = PartitionKeyUtility.GetBatchIdFromBatchPartitionKey(batchPartitionKey);

            if (!context.IsReplaying)
            {
                log.LogInformation($"About to get recipients from batch {batchId}.");
            }

            var recipients = await context.CallActivityAsync<IEnumerable<SentNotificationDataEntity>>(
                FunctionNames.GetRecipientsActivity,
                batchPartitionKey,
                FunctionSettings.DefaultRetryOptions);

            var batches = recipients.AsBatches(SendQueue.MaxNumberOfMessagesInBatchRequest).ToList();

            var totalBatchCount = batches.Count;
            if (!context.IsReplaying)
            {
                log.LogInformation($"About to process {totalBatchCount} batches.");
            }

            var tasks = new List<Task>();
            for (var batchIndex = 0; batchIndex < totalBatchCount; batchIndex++)
            {
                if (!context.IsReplaying)
                {
                    log.LogInformation($"About to process batch {batchIndex + 1} / {totalBatchCount}");
                }

                var task = context.CallActivityAsync(
                    FunctionNames.SendBatchMessagesActivity,
                    (notificationId, batches[batchIndex]),
                    FunctionSettings.DefaultRetryOptions);

                tasks.Add(task);
            }

            await Task.WhenAll(tasks);
        }
    }
}
