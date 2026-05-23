// <copyright file="GraphTokenProvider.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Authentication
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Identity.Web;
    using Microsoft.Kiota.Abstractions.Authentication;

    /// <summary>
    /// Provides access tokens for Microsoft Graph API calls (OBO / delegate flow).
    /// </summary>
    public class GraphTokenProvider : IAccessTokenProvider
    {
        private readonly ITokenAcquisition tokenAcquisition;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphTokenProvider"/> class.
        /// </summary>
        /// <param name="tokenAcquisition">MSAL.NET token acquisition service.</param>
        public GraphTokenProvider(ITokenAcquisition tokenAcquisition)
        {
            this.tokenAcquisition = tokenAcquisition ?? throw new ArgumentNullException(nameof(tokenAcquisition));
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
            // Use OBO flow (delegate) for all Graph calls from the web app.
            return await this.tokenAcquisition.GetAccessTokenForUserAsync(
                new[] { Common.Constants.ScopeDefault });
        }
    }
}
