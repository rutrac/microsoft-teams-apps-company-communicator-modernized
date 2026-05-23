// <copyright file="MockAuthenticationHelper.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.App.CompanyCommunicator.Common.Test.Services.Mock
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Kiota.Abstractions.Authentication;

    /// <summary>
    /// Mocking Authentication Provider (IAccessTokenProvider for Kiota/Graph v5).
    /// </summary>
    public class MockAuthenticationHelper : IAccessTokenProvider
    {
        /// <inheritdoc/>
        public AllowedHostsValidator AllowedHostsValidator { get; } =
            new AllowedHostsValidator(new[] { "graph.microsoft.com" });

        /// <inheritdoc/>
        public Task<string> GetAuthorizationTokenAsync(
            Uri uri,
            Dictionary<string, object>? additionalAuthenticationContext = null,
            CancellationToken cancellationToken = default)
        {
            return Task.FromResult("fake-token");
        }
    }
}
