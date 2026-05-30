// <copyright file="CCBotFrameworkHttpAdapter.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Adapter
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Secrets;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;

    /// <summary>
    /// Bot framework http adapter instance.
    /// </summary>
    public class CCBotFrameworkHttpAdapter : BotFrameworkHttpAdapter, ICCBotFrameworkHttpAdapter
    {
        private readonly ICertificateProvider certificateProvider;
        private readonly BotOptions botOptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="CCBotFrameworkHttpAdapter"/> class.
        /// </summary>
        /// <param name="credentialProvider">credential provider.</param>
        /// <param name="certificateProvider">certificate provider.</param>
        /// <param name="botOptions">bot options (tenantId + per-app passwords).</param>
        public CCBotFrameworkHttpAdapter(
            ICredentialProvider credentialProvider,
            ICertificateProvider certificateProvider,
            IOptions<BotOptions> botOptions)
            : base(credentialProvider)
        {
            this.certificateProvider = certificateProvider;
            this.botOptions = botOptions?.Value ?? throw new ArgumentNullException(nameof(botOptions));
        }

        /// <inheritdoc/>
        public async Task CreateConversationUsingCertificateAsync(string channelId, string serviceUrl, AppCredentials appCredentials, ConversationParameters conversationParameters, BotCallbackHandler callback, CancellationToken cancellationToken)
        {
            var cert = await this.certificateProvider.GetCertificateAsync(appCredentials.MicrosoftAppId);
            var options = new CertificateAppCredentialsOptions()
            {
                AppId = appCredentials.MicrosoftAppId,
                ClientCertificate = cert,
            };

            MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);
            await this.CreateConversationAsync(channelId, serviceUrl, new CertificateAppCredentials(options) as AppCredentials, conversationParameters, callback, cancellationToken);
        }

        /// <inheritdoc/>
        public async Task CreateConversationUsingSecretAsync(string channelId, string serviceUrl, MicrosoftAppCredentials credentials, ConversationParameters conversationParameters, BotCallbackHandler callback, CancellationToken cancellationToken)
        {
            MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);
            await this.CreateConversationAsync(channelId, serviceUrl, credentials, conversationParameters, callback, cancellationToken);
        }

        /// <inheritdoc/>
        protected override async Task<AppCredentials> BuildCredentialsAsync(string appId, string oAuthScope = null)
        {
            appId = appId ?? throw new ArgumentNullException(nameof(appId));

            if (this.certificateProvider.IsCertificateAuthenticationEnabled())
            {
                var cert = await this.certificateProvider.GetCertificateAsync(appId);
                var options = new CertificateAppCredentialsOptions()
                {
                    AppId = appId,
                    ClientCertificate = cert,
                    OauthScope = oAuthScope,
                };

                var certificateAppCredentials = new CertificateAppCredentials(options) as AppCredentials;
                return certificateAppCredentials;
            }
            else
        {
                // SingleTenant bot apps require channelAuthTenant so the SDK hits the tenant-specific
                // /oauth2/token endpoint instead of the default /botframework.com one. Resolve the
                // password from BotOptions (the registered ICredentialProvider only knows the legacy
                // MicrosoftAppId/MicrosoftAppPassword pair and returns empty for our UserAppId/AuthorAppId).
                var tenantId = string.IsNullOrWhiteSpace(this.botOptions.TenantId) ? null : this.botOptions.TenantId;
                string password = null;
                if (!string.IsNullOrEmpty(this.botOptions.UserAppId) && string.Equals(appId, this.botOptions.UserAppId, StringComparison.OrdinalIgnoreCase))
                {
                    password = this.botOptions.UserAppPassword;
                }
                else if (!string.IsNullOrEmpty(this.botOptions.AuthorAppId) && string.Equals(appId, this.botOptions.AuthorAppId, StringComparison.OrdinalIgnoreCase))
                {
                    password = this.botOptions.AuthorAppPassword;
                }
                else
                {
                    password = await this.CredentialProvider.GetAppPasswordAsync(appId);
                }

                return new MicrosoftAppCredentials(appId, password, channelAuthTenant: tenantId, oAuthScope: oAuthScope);
            }
        }
    }
}
