// <copyright file="StorageClientFactory.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Clients
{
    using System;
    using global::Azure.Core;
    using global::Azure.Identity;
    using global::Azure.Storage.Blobs;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;

    /// <summary>
    /// Storage client factory.
    /// </summary>
    public class StorageClientFactory : IStorageClientFactory
    {
        private readonly string storageConnectionString;
        private readonly string storageAccountName;
        private readonly bool useManagedIdentity;

        /// <summary>
        /// Initializes a new instance of the <see cref="StorageClientFactory"/> class.
        /// </summary>
        /// <param name="repositoryOptions">Repository options.</param>
        /// <param name="configuration">App configuration (for UseManagedIdentity + StorageAccountName).</param>
        public StorageClientFactory(IOptions<RepositoryOptions> repositoryOptions, IConfiguration configuration)
        {
            this.storageConnectionString = repositoryOptions.Value.StorageAccountConnectionString;
            this.storageAccountName = configuration.GetValue<string>("StorageAccountName");
            this.useManagedIdentity = configuration.GetValue<bool>("UseManagedIdentity");
        }

        /// <inheritdoc/>
        public BlobContainerClient CreateBlobContainerClient()
        {
            return this.CreateBlobContainerClient(Constants.BlobContainerName);
        }

        /// <inheritdoc/>
        public BlobContainerClient CreateBlobContainerClient(string blobContainerName)
        {
            var options = new BlobClientOptions();
            options.Retry.MaxRetries = 5;
            options.Retry.Mode = RetryMode.Exponential;
            options.Retry.Delay = TimeSpan.FromSeconds(1);

            if (this.useManagedIdentity)
            {
                var blobContainerUri = new Uri($"https://{this.storageAccountName}.blob.core.windows.net/{blobContainerName}");
                return new BlobContainerClient(blobContainerUri, new DefaultAzureCredential(), options);
            }

            return new BlobContainerClient(this.storageConnectionString, blobContainerName, options);
        }
    }
}
