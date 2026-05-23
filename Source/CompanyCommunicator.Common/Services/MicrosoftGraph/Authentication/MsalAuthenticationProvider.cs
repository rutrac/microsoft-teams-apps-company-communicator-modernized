// <copyright file="MsalAuthenticationProvider.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Identity.Client;
    using Microsoft.Kiota.Abstractions.Authentication;

    /// <summary>
    /// MSAL authentication provider for Graph calls (client credentials / app-only flow).
    /// </summary>
    public class MsalAuthenticationProvider : IAccessTokenProvider
    {
        private readonly IConfidentialClientApplication clientApplication;

        /// <summary>
        /// Initializes a new instance of the <see cref="MsalAuthenticationProvider"/> class.
        /// </summary>
        /// <param name="clientApplication">MSAL.NET confidential client application.</param>
        public MsalAuthenticationProvider(IConfidentialClientApplication clientApplication)
        {
            this.clientApplication = clientApplication ?? throw new ArgumentNullException(nameof(clientApplication));
        }

        /// <inheritdoc/>
        public AllowedHostsValidator AllowedHostsValidator { get; } =
            new AllowedHostsValidator(new[] { "graph.microsoft.com" });

        /// <inheritdoc/>
        public async Task<string> GetAuthorizationTokenAsync(
            Uri uri,
            Dictionary<string, object>? additionalAuthenticationContext = null,
            CancellationToken cancellationToken = default)
        {
            var scopes = new List<string> { Common.Constants.ScopeDefault };
            var result = await this.clientApplication
                .AcquireTokenForClient(scopes)
                .ExecuteAsync(cancellationToken);
            return result.AccessToken;
        }
    }
}
