// <copyright file="DraftNotificationPreviewService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.DraftNotificationPreview
{
    using System;
    using System.Net;
    using System.Threading;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Bot;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;

    /// <summary>
    /// Draft notification preview service.
    /// </summary>
    public class DraftNotificationPreviewService : IDraftNotificationPreviewService
    {
        private static readonly string MsTeamsChannelId = "msteams";
        private static readonly string ChannelConversationType = "channel";
        private static readonly string ThrottledErrorResponse = "Throttled";

        private readonly string botAppId;
        private readonly AdaptiveCardCreator adaptiveCardCreator;
        private readonly CompanyCommunicatorBotAdapter companyCommunicatorBotAdapter;
        private readonly ILogger<DraftNotificationPreviewService> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="DraftNotificationPreviewService"/> class.
        /// </summary>
        /// <param name="botOptions">The bot options.</param>
        /// <param name="adaptiveCardCreator">Adaptive card creator service.</param>
        /// <param name="companyCommunicatorBotAdapter">Bot framework http adapter instance.</param>
        /// <param name="logger">Logger.</param>
        public DraftNotificationPreviewService(
            IOptions<BotOptions> botOptions,
            AdaptiveCardCreator adaptiveCardCreator,
            CompanyCommunicatorBotAdapter companyCommunicatorBotAdapter,
            ILogger<DraftNotificationPreviewService> logger)
        {
            this.logger = logger ?? throw new ArgumentNullException(nameof(logger));
            var options = botOptions ?? throw new ArgumentNullException(nameof(botOptions));

            // Preview posts into the destination Teams channel using the User bot, because the
            // User app is the one installed in destination teams (the Authors app lives only in
            // an authoring team and would not have a TeamData/ServiceUrl row for arbitrary teams).
            this.botAppId = options.Value.UserAppId;
            if (string.IsNullOrEmpty(this.botAppId))
            {
                throw new ApplicationException("UserAppId setting is missing in the configuration.");
            }

            this.adaptiveCardCreator = adaptiveCardCreator ?? throw new ArgumentNullException(nameof(adaptiveCardCreator));
            this.companyCommunicatorBotAdapter = companyCommunicatorBotAdapter ?? throw new ArgumentNullException(nameof(companyCommunicatorBotAdapter));
        }

        /// <inheritdoc/>
        public async Task<HttpStatusCode> SendPreview(NotificationDataEntity draftNotificationEntity, TeamDataEntity teamDataEntity, string teamsChannelId)
        {
            if (draftNotificationEntity == null)
            {
                throw new ArgumentException("Null draft notification entity.");
            }

            if (teamDataEntity == null)
            {
                throw new ArgumentException("Null team data entity.");
            }

            if (string.IsNullOrWhiteSpace(teamsChannelId))
            {
                throw new ArgumentException("Null channel id.");
            }

            if (string.IsNullOrWhiteSpace(teamDataEntity.ServiceUrl) || string.IsNullOrWhiteSpace(teamDataEntity.TenantId))
            {
                // Without a real service URL + tenant id the bot adapter will return 502 from the channel.
                return HttpStatusCode.ServiceUnavailable;
            }

            // Trust the Teams service URL so the Bot Framework will issue tokens for it.
            Microsoft.Bot.Connector.Authentication.MicrosoftAppCredentials.TrustServiceUrl(teamDataEntity.ServiceUrl);
            this.logger.LogWarning("SENDPREVIEW step=trusted-service-url botAppId={BotAppId} serviceUrl={ServiceUrl} tenantId={TenantId} channelId={ChannelId}", this.botAppId, teamDataEntity.ServiceUrl, teamDataEntity.TenantId, teamsChannelId);

            // Create bot conversation reference.
            var conversationReference = this.PrepareConversationReferenceAsync(teamDataEntity, teamsChannelId);
            this.logger.LogWarning("SENDPREVIEW step=ref-built");

            // Trigger bot to send the adaptive card.
            try
            {
                this.logger.LogWarning("SENDPREVIEW step=before-continueconversation");
                await this.companyCommunicatorBotAdapter.ContinueConversationAsync(
                    this.botAppId,
                    conversationReference,
                    async (turnContext, cancellationToken) =>
                    {
                        this.logger.LogWarning("SENDPREVIEW step=callback-entered activityType={Type} channelId={Channel}", turnContext?.Activity?.Type, turnContext?.Activity?.ChannelId);
                        try
                        {
                            await this.SendAdaptiveCardAsync(turnContext, draftNotificationEntity);
                            this.logger.LogWarning("SENDPREVIEW step=callback-sendactivity-completed");
                        }
                        catch (Exception cbEx)
                        {
                            this.logger.LogError(cbEx, "SENDPREVIEW step=callback-exception type={Type} msg={Msg}", cbEx.GetType().FullName, cbEx.Message);
                            throw;
                        }
                    },
                    CancellationToken.None);
                this.logger.LogWarning("SENDPREVIEW step=continueconversation-returned");
                return HttpStatusCode.OK;
            }
            catch (ErrorResponseException e)
            {
                this.logger.LogError(e, "SENDPREVIEW step=errorresponse code={Code} msg={Msg}", e.Body?.Error?.Code, e.Message);
                var errorResponse = (ErrorResponse)e.Body;
                if (errorResponse != null
                    && errorResponse.Error.Code.Equals(DraftNotificationPreviewService.ThrottledErrorResponse, StringComparison.OrdinalIgnoreCase))
                {
                    return HttpStatusCode.TooManyRequests;
                }

                throw;
            }
        }

        private ConversationReference PrepareConversationReferenceAsync(TeamDataEntity teamDataEntity, string channelId)
        {
            var channelAccount = new ChannelAccount
            {
                Id = $"28:{this.botAppId}",
            };

            var conversationAccount = new ConversationAccount
            {
                ConversationType = DraftNotificationPreviewService.ChannelConversationType,
                Id = channelId,
                TenantId = teamDataEntity.TenantId,
            };

            var conversationReference = new ConversationReference
            {
                Bot = channelAccount,
                ChannelId = DraftNotificationPreviewService.MsTeamsChannelId,
                Conversation = conversationAccount,
                ServiceUrl = teamDataEntity.ServiceUrl,
            };

            return conversationReference;
        }

        private async Task SendAdaptiveCardAsync(
            ITurnContext turnContext,
            NotificationDataEntity draftNotificationEntity)
        {
            WriteBreadcrumb("send-enter");
            var reply = this.CreateReply(draftNotificationEntity);
            WriteBreadcrumb("send-reply-built");
            await turnContext.SendActivityAsync(reply);
            WriteBreadcrumb("send-completed");
        }

        private static void WriteBreadcrumb(string step)
        {
            try
            {
                var dir = System.IO.Path.Combine(
                    System.Environment.GetEnvironmentVariable("HOME") ?? "D:\\home",
                    "LogFiles",
                    "preview-breadcrumbs");
                System.IO.Directory.CreateDirectory(dir);
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(dir, $"breadcrumbs-{System.DateTime.UtcNow:yyyyMMdd}.log"),
                    $"{System.DateTime.UtcNow:O} pid={System.Environment.ProcessId} tid={System.Environment.CurrentManagedThreadId} step={step}\n");
            }
            catch
            {
                // best-effort
            }
        }

        private IMessageActivity CreateReply(NotificationDataEntity draftNotificationEntity)
        {
            var adaptiveCard = this.adaptiveCardCreator.CreateAdaptiveCard(
                draftNotificationEntity.Title,
                draftNotificationEntity.ImageLink,
                draftNotificationEntity.Summary,
                draftNotificationEntity.Author,
                draftNotificationEntity.ButtonTitle,
                draftNotificationEntity.ButtonLink,
                draftNotificationEntity.Buttons,
                string.Empty,
                draftNotificationEntity.ChannelImage,
                draftNotificationEntity.ChannelTitle);

            var attachment = new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard,
            };

            var reply = MessageFactory.Attachment(attachment);

            return reply;
        }
    }
}