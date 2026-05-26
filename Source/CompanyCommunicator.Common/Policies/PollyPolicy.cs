// <copyright file="PollyPolicy.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Policies
{
    using System;
    using System.Net;
    using Microsoft.Graph.Models.ODataErrors;
    using Polly;
    using Polly.Retry;

    /// <summary>
    /// Polly policies.
    /// </summary>
    public class PollyPolicy
    {
        /// <summary>
        /// Get the graph retry pipeline.
        /// </summary>
        /// <param name="maxAttempts">the number of max attempts.</param>
        /// <returns>A resilience pipeline that can be applied to async delegates.</returns>
        public static ResiliencePipeline GetGraphRetryPolicy(int maxAttempts)
        {
            // Only Handling 502 Bad Gateway Exception
            // Other exception such as 429, 503, 504 is handled by default by Graph SDK.
            return new ResiliencePipelineBuilder()
                .AddRetry(new RetryStrategyOptions
                {
                    ShouldHandle = new PredicateBuilder().Handle<ODataError>(e =>
                        e.ResponseStatusCode == (int)HttpStatusCode.BadGateway),
                    MaxRetryAttempts = maxAttempts,
                    BackoffType = DelayBackoffType.Exponential,
                    UseJitter = true,
                    Delay = TimeSpan.FromSeconds(1),
                })
                .Build();
        }
    }
}
