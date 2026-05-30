// <copyright file="CompanyCommunicatorBotAdapter.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Bot
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Secrets;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;

    /// <summary>
    /// The Company Communicator Bot Adapter.
    /// </summary>
    public class CompanyCommunicatorBotAdapter : BotFrameworkHttpAdapter
    {
        private readonly ICertificateProvider certificateProvider;
        private readonly BotOptions botOptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyCommunicatorBotAdapter"/> class.
        /// </summary>
        /// <param name="credentialProvider">Credential provider service instance.</param>
        /// <param name="companyCommunicatorBotFilterMiddleware">Teams message filter middleware instance.</param>
        /// <param name="certificateProvider">Certificate provider service instance.</param>
        /// <param name="botOptions">Bot options (tenantId + per-app passwords).</param>
        public CompanyCommunicatorBotAdapter(
            ICredentialProvider credentialProvider,
            CompanyCommunicatorBotFilterMiddleware companyCommunicatorBotFilterMiddleware,
            ICertificateProvider certificateProvider,
            IOptions<BotOptions> botOptions)
            : base(credentialProvider)
        {
            companyCommunicatorBotFilterMiddleware = companyCommunicatorBotFilterMiddleware ?? throw new ArgumentNullException(nameof(companyCommunicatorBotFilterMiddleware));
            this.certificateProvider = certificateProvider ?? throw new ArgumentNullException(nameof(certificateProvider));
            this.botOptions = botOptions?.Value ?? throw new ArgumentNullException(nameof(botOptions));

            // Middleware
            this.Use(companyCommunicatorBotFilterMiddleware);
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

                return new CertificateAppCredentials(options) as AppCredentials;
            }

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
