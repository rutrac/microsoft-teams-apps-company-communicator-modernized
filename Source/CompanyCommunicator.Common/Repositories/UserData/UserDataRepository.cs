// <copyright file="UserDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>
namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData
{
    using System;
    using System.Threading.Tasks;
    using Azure.Data.Tables;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;

    /// <summary>
    /// Repository of the user data stored in the table storage.
    /// </summary>
    public class UserDataRepository : BaseRepository<UserDataEntity>, IUserDataRepository
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UserDataRepository"/> class.
        /// </summary>
        /// <param name="logger">The logging service.</param>
        /// <param name="repositoryOptions">Options used to create the repository.</param>
        public UserDataRepository(
            ILogger<UserDataRepository> logger,
            IOptions<RepositoryOptions> repositoryOptions)
            : base(
                  logger,
                  storageAccountName: repositoryOptions.Value.StorageAccountName,
                  tableName: UserDataTableNames.TableName,
                  defaultPartitionKey: UserDataTableNames.UserDataPartition,
                  ensureTableExists: repositoryOptions.Value.EnsureTableExists)
        {
        }

        /// <inheritdoc/>
        public async Task<string> GetDeltaLinkAsync()
        {
            try
            {
                var response = await this.Table.GetEntityAsync<UsersSyncEntity>(
                    UserDataTableNames.UsersSyncDataPartition,
                    UserDataTableNames.AllUsersDeltaLinkRowKey);
                return response.Value?.Value;
            }
            catch (Azure.RequestFailedException ex) when (ex.Status == 404)
            {
                return null;
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task SetDeltaLinkAsync(string deltaLink)
        {
            if (string.IsNullOrEmpty(deltaLink))
            {
                throw new ArgumentNullException(nameof(deltaLink));
            }

            var entity = new UsersSyncEntity()
            {
                PartitionKey = UserDataTableNames.UsersSyncDataPartition,
                RowKey = UserDataTableNames.AllUsersDeltaLinkRowKey,
                Value = deltaLink,
            };

            try
            {
                await this.Table.UpsertEntityAsync(entity, TableUpdateMode.Replace);
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }
    }
}
