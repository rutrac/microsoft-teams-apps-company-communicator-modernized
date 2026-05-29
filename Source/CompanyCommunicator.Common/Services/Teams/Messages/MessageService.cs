// <copyright file="MessageService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams
{
    using System;
    using System.Net;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Adapter;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;
    using Polly;
    using Polly.Retry;

    /// <summary>
    /// Teams message service.
    /// </summary>
    public class MessageService : IMessageService
    {
        private readonly string microsoftAppId;
        private readonly ICCBotFrameworkHttpAdapter botAdapter;

        /// <summary>
        /// Initializes a new instance of the <see cref="MessageService"/> class.
        /// </summary>
        /// <param name="botOptions">The bot options.</param>
        /// <param name="botAdapter">The bot adapter.</param>
        public MessageService(
            IOptions<BotOptions> botOptions,
            ICCBotFrameworkHttpAdapter botAdapter)
        {
            this.microsoftAppId = botOptions?.Value?.UserAppId ?? throw new ArgumentNullException(nameof(botOptions));
            this.botAdapter = botAdapter ?? throw new ArgumentNullException(nameof(botAdapter));
        }

        /// <inheritdoc/>
        public async Task<SendMessageResponse> SendMessageAsync(
            IMessageActivity message,
            string conversationId,
            string serviceUrl,
            int maxAttempts,
            ILogger log)
        {
            if (message is null)
            {
                throw new ArgumentNullException(nameof(message));
            }

            if (string.IsNullOrEmpty(conversationId))
            {
                throw new ArgumentException($"'{nameof(conversationId)}' cannot be null or empty", nameof(conversationId));
            }

            if (string.IsNullOrEmpty(serviceUrl))
            {
                throw new ArgumentException($"'{nameof(serviceUrl)}' cannot be null or empty", nameof(serviceUrl));
            }

            if (log is null)
            {
                throw new ArgumentNullException(nameof(log));
            }

            var conversationReference = new ConversationReference
            {
                ServiceUrl = serviceUrl,
                Conversation = new ConversationAccount
                {
                    Id = conversationId,
                },
            };

            var response = new SendMessageResponse
            {
                TotalNumberOfSendThrottles = 0,
                AllSendStatusCodes = string.Empty,
            };

            // Trust the Teams service URL so the Bot Framework will issue tokens for it.
            // Without this, proactive sends can be silently dropped on the first use of a new service URL.
            MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);

            await this.botAdapter.ContinueConversationAsync(
                botId: this.microsoftAppId,
                reference: conversationReference,
                callback: async (turnContext, cancellationToken) =>
                {
                    var policy = this.GetRetryPolicy(maxAttempts, log);
                    try
                    {
                        await policy.ExecuteAsync(async ct =>
                        {
                            var resp = await turnContext.SendActivityAsync(message, ct);

                            // A successful Teams delivery always returns a non-empty activity id.
                            // A null/empty id means the call was accepted upstream but Teams did not
                            // actually deliver the message (e.g. stale conversation, untrusted URL).
                            if (resp == null || string.IsNullOrEmpty(resp.Id))
                            {
                                throw new InvalidOperationException("Bot send returned no activity id; treating as not delivered.");
                            }

                            response.ActivityId = resp.Id;
                        });

                        // Success.
                        response.ResultType = SendMessageResult.Succeeded;
                        response.StatusCode = (int)HttpStatusCode.Created;
                        response.AllSendStatusCodes += $"{(int)HttpStatusCode.Created},";
                    }
                    catch (ErrorResponseException exception)
                    {
                        var errorMessage = $"{exception.GetType()}: {exception.Message}";
                        log.LogError(exception, $"Failed to send message. Exception message: {errorMessage}");

                        response.StatusCode = (int)exception.Response.StatusCode;
                        response.AllSendStatusCodes += $"{(int)exception.Response.StatusCode},";
                        response.ErrorMessage = exception.ToString();
                        switch (exception.Response.StatusCode)
                        {
                            case HttpStatusCode.TooManyRequests:
                                response.ResultType = SendMessageResult.Throttled;
                                response.TotalNumberOfSendThrottles = maxAttempts;
                                break;

                            case HttpStatusCode.NotFound:
                                response.ResultType = SendMessageResult.RecipientNotFound;
                                break;

                            default:
                                response.ResultType = SendMessageResult.Failed;
                                break;
                        }
                    }
                    catch (InvalidOperationException exception)
                    {
                        log.LogError(exception, $"Failed to send message: {exception.Message}");
                        response.StatusCode = (int)HttpStatusCode.BadGateway;
                        response.AllSendStatusCodes += $"{(int)HttpStatusCode.BadGateway},";
                        response.ErrorMessage = exception.ToString();
                        response.ResultType = SendMessageResult.Failed;
                    }
                },
                cancellationToken: CancellationToken.None);

            return response;
        }

        private ResiliencePipeline GetRetryPolicy(int maxAttempts, ILogger log)
        {
            return new ResiliencePipelineBuilder()
                .AddRetry(new RetryStrategyOptions
                {
                    ShouldHandle = new PredicateBuilder().Handle<ErrorResponseException>(e =>
                    {
                        var errorMessage = $"{e.GetType()}: {e.Message}";
                        log.LogError(e, $"Exception thrown: {errorMessage}");

                        var statusCode = e.Response.StatusCode;

                        // Auth failures (401/403) will not recover by retrying; fail fast.
                        if (statusCode == HttpStatusCode.Unauthorized || statusCode == HttpStatusCode.Forbidden)
                        {
                            return false;
                        }

                        // Handle throttling and transient server errors.
                        return statusCode == HttpStatusCode.TooManyRequests || ((int)statusCode >= 500 && (int)statusCode < 600);
                    }),
                    MaxRetryAttempts = maxAttempts,
                    BackoffType = DelayBackoffType.Exponential,
                    UseJitter = true,
                    Delay = TimeSpan.FromSeconds(5),
                    MaxDelay = TimeSpan.FromSeconds(60),
                })
                .Build();
        }
    }
}
