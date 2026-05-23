// <copyright file="MockHttpProvider.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.App.CompanyCommunicator.Common.Test.Services.Mock
{
    using System;
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Newtonsoft.Json;

    /// <summary>
    /// Mocking Http message handler for Graph SDK v5 tests.
    /// </summary>
    public class MockHttpProvider : HttpMessageHandler
    {
        /// <summary>
        /// Gets or sets response mapping with key, response.
        /// </summary>
        public Dictionary<string, object> Responses { get; set; } = new Dictionary<string, object>();

        /// <inheritdoc/>
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            string key = request.Method.ToString() + ":" + request.RequestUri.ToString();
            var response = new HttpResponseMessage(System.Net.HttpStatusCode.OK);
            if (this.Responses.TryGetValue(key, out var body))
            {
                response.Content = new StringContent(
                    JsonConvert.SerializeObject(body),
                    Encoding.UTF8,
                    "application/json");
            }

            return Task.FromResult(response);
        }
    }
}
