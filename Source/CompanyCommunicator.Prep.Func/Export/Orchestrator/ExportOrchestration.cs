// <copyright file="ExportOrchestration.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.Export.Orchestrator
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ExportData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.Export.Model;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend;

    /// <summary>
    /// Durable orchestration for exporting notifications.
    /// </summary>
    public static class ExportOrchestration
    {
        /// <summary>
        /// Run orchestrator.
        /// </summary>
        /// <param name="context">Durable orchestration context.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.ExportOrchestration)]
        public static async Task RunOrchestrator(
            [OrchestrationTrigger] TaskOrchestrationContext context)
        {
            var log = context.CreateReplaySafeLogger(nameof(ExportOrchestration));
            var exportRequiredData = context.GetInput<ExportDataRequirement>();
            var sentNotificationDataEntity = exportRequiredData.NotificationDataEntity;
            var exportDataEntity = exportRequiredData.ExportDataEntity;

            if (!context.IsReplaying)
            {
                log.LogInformation($"Start to export the notification {sentNotificationDataEntity.Id}!");
            }

            try
            {
                if (!context.IsReplaying)
                {
                    log.LogInformation("About to update export is in progress.");
                }

                exportDataEntity.Status = ExportStatus.InProgress.ToString();
                await context.CallActivityAsync(
                    FunctionNames.UpdateExportDataActivity,
                    exportDataEntity,
                    FunctionSettings.DefaultRetryOptions);

                if (!context.IsReplaying)
                {
                    log.LogInformation("About to get the metadata information.");
                }

                var metaData = await context.CallActivityAsync<Metadata>(
                    FunctionNames.GetMetadataActivity,
                    (sentNotificationDataEntity, exportDataEntity),
                    FunctionSettings.DefaultRetryOptions);

                if (!context.IsReplaying)
                {
                    log.LogInformation("About to start file upload.");
                }

                await context.CallActivityAsync(
                    FunctionNames.UploadActivity,
                    (sentNotificationDataEntity, metaData, exportDataEntity.FileName),
                    FunctionSettings.DefaultRetryOptions);

                if (!context.IsReplaying)
                {
                    log.LogInformation("About to send file card.");
                }

                var consentId = await context.CallActivityAsync<string>(
                    FunctionNames.SendFileCardActivity,
                    (exportRequiredData.UserId, exportRequiredData.NotificationDataEntity.Id, exportDataEntity.FileName),
                    FunctionSettings.DefaultRetryOptions);

                if (!context.IsReplaying)
                {
                    log.LogInformation("About to update export is completed.");
                }

                exportDataEntity.FileConsentId = consentId;
                exportDataEntity.Status = ExportStatus.Completed.ToString();
                await context.CallActivityAsync(
                    FunctionNames.UpdateExportDataActivity,
                    exportDataEntity,
                    FunctionSettings.DefaultRetryOptions);

                log.LogInformation($"ExportOrchestration is successful for notification id:{sentNotificationDataEntity.Id}!");
            }
            catch (Exception ex)
            {
                log.LogError(ex, $"Failed to export notification {sentNotificationDataEntity.Id} : {ex.Message}");
                await context.CallActivityAsync(
                    FunctionNames.HandleExportFailureActivity,
                    exportDataEntity,
                    FunctionSettings.DefaultRetryOptions);
            }
        }
    }
}
