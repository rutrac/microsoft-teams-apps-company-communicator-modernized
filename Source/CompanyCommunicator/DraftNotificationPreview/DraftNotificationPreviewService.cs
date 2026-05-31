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
        private readonly BotOptions botOptions;
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

            this.botOptions = options.Value;
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
                WriteBreadcrumb("before-continueconversation");
                using var cts = new System.Threading.CancellationTokenSource(System.TimeSpan.FromSeconds(30));
                var ccTask = this.companyCommunicatorBotAdapter.ContinueConversationAsync(
                    this.botAppId,
                    conversationReference,
                    async (turnContext, cancellationToken) =>
                    {
                        WriteBreadcrumb("callback-entered");
                        this.logger.LogWarning("SENDPREVIEW step=callback-entered activityType={Type} channelId={Channel}", turnContext?.Activity?.Type, turnContext?.Activity?.ChannelId);
                        try
                        {
                            await this.SendAdaptiveCardAsync(turnContext, draftNotificationEntity);
                            WriteBreadcrumb("callback-sendactivity-completed");
                            this.logger.LogWarning("SENDPREVIEW step=callback-sendactivity-completed");
                        }
                        catch (Exception cbEx)
                        {
                            WriteBreadcrumb($"callback-exception type={cbEx.GetType().FullName} msg={cbEx.Message} inner={cbEx.InnerException?.GetType().FullName}:{cbEx.InnerException?.Message}\nSTACK:\n{cbEx}");
                            this.logger.LogError(cbEx, "SENDPREVIEW step=callback-exception type={Type} msg={Msg}", cbEx.GetType().FullName, cbEx.Message);
                            throw;
                        }
                    },
                    cts.Token);
                var winner = await System.Threading.Tasks.Task.WhenAny(ccTask, System.Threading.Tasks.Task.Delay(System.TimeSpan.FromSeconds(35)));
                if (winner != ccTask)
                {
                    WriteBreadcrumb("TIMEOUT-35s waiting for ContinueConversationAsync");
                    this.logger.LogError("SENDPREVIEW step=hang-timeout - ContinueConversationAsync did not complete in 35s; returning 504");
                    return HttpStatusCode.GatewayTimeout;
                }

                await ccTask; // surface any exception
                WriteBreadcrumb("continueconversation-returned");
                this.logger.LogWarning("SENDPREVIEW step=continueconversation-returned");
                return HttpStatusCode.OK;
            }
            catch (ErrorResponseException e)
            {
                WriteBreadcrumb($"errorresponse code={e.Body?.Error?.Code} msg={e.Message}\nSTACK:\n{e}");
                this.logger.LogError(e, "SENDPREVIEW step=errorresponse code={Code} msg={Msg}", e.Body?.Error?.Code, e.Message);
                var errorResponse = (ErrorResponse)e.Body;
                if (errorResponse != null
                    && errorResponse.Error.Code.Equals(DraftNotificationPreviewService.ThrottledErrorResponse, StringComparison.OrdinalIgnoreCase))
                {
                    return HttpStatusCode.TooManyRequests;
                }

                throw;
            }
            catch (Exception ex)
            {
                WriteBreadcrumb($"outer-exception type={ex.GetType().FullName} msg={ex.Message} inner={ex.InnerException?.GetType().FullName}:{ex.InnerException?.Message}\nSTACK:\n{ex}");
                this.logger.LogError(ex, "SENDPREVIEW step=outer-exception type={Type} msg={Msg}", ex.GetType().FullName, ex.Message);
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

            // Direct SDK credential + token probe: builds the same MicrosoftAppCredentials the
            // adapter will use, then forces a token acquisition. If the worker dies in here, the
            // crash is in MSAL/credential code (not the conversations REST call).
            try
            {
                var tenantId = string.IsNullOrWhiteSpace(this.botOptions.TenantId) ? null : this.botOptions.TenantId;
                WriteBreadcrumb($"probe-cred-pre tenant={tenantId} pwdLen={this.botOptions.UserAppPassword?.Length ?? 0}");
                var directCreds = new Microsoft.Bot.Connector.Authentication.MicrosoftAppCredentials(
                    this.botAppId,
                    this.botOptions.UserAppPassword,
                    channelAuthTenant: tenantId);
                WriteBreadcrumb("probe-cred-built");
                var token = await directCreds.GetTokenAsync();
                WriteBreadcrumb($"probe-token-acquired len={token?.Length ?? 0}");
            }
            catch (System.Exception probeEx)
            {
                WriteBreadcrumb($"probe-token-exception type={probeEx.GetType().FullName} msg={probeEx.Message} inner={probeEx.InnerException?.GetType().FullName}:{probeEx.InnerException?.Message}");
            }

            WriteBreadcrumb("send-before-sendactivity");

            // BISECT: send a plain text activity first. If this succeeds, the AdaptiveCard
            // attachment serialization is what kills the SDK pipeline.
            try
            {
                var textOnly = Microsoft.Bot.Builder.MessageFactory.Text("preview diag: plain text probe");
                WriteBreadcrumb("text-before-send");
                var textTask = turnContext.SendActivityAsync(textOnly);
                WriteBreadcrumb($"text-task-created status={textTask.Status}");
                var textWinner = await System.Threading.Tasks.Task.WhenAny(textTask, System.Threading.Tasks.Task.Delay(System.TimeSpan.FromSeconds(15)));
                if (textWinner != textTask)
                {
                    WriteBreadcrumb($"text-TIMEOUT-15s status={textTask.Status}");
                }
                else
                {
                    var tr = await textTask;
                    WriteBreadcrumb($"text-completed respId={tr?.Id}");
                }
            }
            catch (System.Exception textEx)
            {
                WriteBreadcrumb($"text-exception type={textEx.GetType().FullName} msg={textEx.Message}");
            }

            WriteBreadcrumb("card-before-sendactivity");
            var sendTask = turnContext.SendActivityAsync(reply);
            WriteBreadcrumb($"send-task-created status={sendTask.Status}");
            var sendWinner = await System.Threading.Tasks.Task.WhenAny(sendTask, System.Threading.Tasks.Task.Delay(System.TimeSpan.FromSeconds(20)));
            if (sendWinner != sendTask)
            {
                WriteBreadcrumb($"send-TIMEOUT-20s status={sendTask.Status}");
                throw new System.TimeoutException("SendActivityAsync did not complete in 20s");
            }

            try
            {
                var resp = await sendTask;
                WriteBreadcrumb($"send-completed respId={resp?.Id}");
            }
            catch (System.Exception sendEx)
            {
                WriteBreadcrumb($"send-exception type={sendEx.GetType().FullName} msg={sendEx.Message} inner={sendEx.InnerException?.GetType().FullName}:{sendEx.InnerException?.Message}\nSTACK:\n{sendEx}");
                throw;
            }
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
                Content = Newtonsoft.Json.Linq.JObject.Parse(adaptiveCard.ToJson()),
            };

            var reply = MessageFactory.Attachment(attachment);

            return reply;
        }
    }
}