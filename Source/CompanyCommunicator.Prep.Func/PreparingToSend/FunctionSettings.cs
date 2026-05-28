// <copyright file="FunctionSettings.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using Microsoft.DurableTask;

    /// <summary>
    /// Function settings.
    /// </summary>
    public class FunctionSettings
    {
        /// <summary>
        /// A default setting for the retry options for starting an activity / sub-orchestrator.
        /// </summary>
        public static readonly TaskOptions DefaultRetryOptions
            = TaskOptions.FromRetryPolicy(new RetryPolicy(
                maxNumberOfAttempts: 3,
                firstRetryInterval: TimeSpan.FromSeconds(5)));
    }
}
