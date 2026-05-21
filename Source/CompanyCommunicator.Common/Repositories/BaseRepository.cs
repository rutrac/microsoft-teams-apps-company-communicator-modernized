// <copyright file="BaseRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using global::Azure;
    using global::Azure.Data.Tables;
    using global::Azure.Identity;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// Base repository for the data stored in the Azure Table Storage.
    /// </summary>
    /// <typeparam name="T">Entity class type.</typeparam>
    public abstract class BaseRepository<T> : IRepository<T>
        where T : class, ITableEntity, new()
    {
        /// <summary>
        /// Maximum length of error and warning messages to save in the entity.
        /// This limit ensures that we don't hit the Azure table storage limits for the max size of the data
        /// in a column, and the total size of an entity.
        /// </summary>
        public const int MaxMessageLengthToSave = 1024;

        private readonly string defaultPartitionKey;

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseRepository{T}"/> class.
        /// </summary>
        /// <param name="logger">The logging service.</param>
        /// <param name="storageAccountName">The storage account name (used with managed identity).</param>
        /// <param name="tableName">The name of the table in Azure Table Storage.</param>
        /// <param name="defaultPartitionKey">Default partition key value.</param>
        /// <param name="ensureTableExists">Flag to ensure the table is created if it doesn't exist.</param>
        public BaseRepository(
            ILogger logger,
            string storageAccountName,
            string tableName,
            string defaultPartitionKey,
            bool ensureTableExists)
        {
            this.Logger = logger;

            var serviceClient = new TableServiceClient(
                new Uri($"https://{storageAccountName}.table.core.windows.net"),
                new DefaultAzureCredential());
            this.Table = serviceClient.GetTableClient(tableName);
            this.defaultPartitionKey = defaultPartitionKey;

            if (ensureTableExists)
            {
                this.Table.CreateIfNotExists();
            }
        }

        /// <inheritdoc/>
        public TableClient Table { get; }

        /// <summary>
        /// Gets the logger service.
        /// </summary>
        protected ILogger Logger { get; }

        /// <inheritdoc/>
        public virtual async Task CreateOrUpdateAsync(T entity)
        {
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

        /// <inheritdoc/>
        public async Task InsertOrMergeAsync(T entity)
        {
            try
            {
                await this.Table.UpsertEntityAsync(entity, TableUpdateMode.Merge);
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task DeleteAsync(T entity)
        {
            try
            {
                await this.Table.DeleteEntityAsync(entity.PartitionKey, entity.RowKey, ETag.All);
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public virtual async Task<T> GetAsync(string partitionKey, string rowKey)
        {
            try
            {
                var response = await this.Table.GetEntityAsync<T>(partitionKey, rowKey);
                return response.Value;
            }
            catch (RequestFailedException ex) when (ex.Status == 404)
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
        public async Task<IEnumerable<T>> GetWithFilterAsync(string filter, string partition = null, int? count = null)
        {
            try
            {
                var partitionKeyFilter = this.GetPartitionKeyFilter(partition);
                var combinedFilter = this.CombineFilters(filter, partitionKeyFilter);
                var entities = await this.ExecuteQueryAsync(combinedFilter, count);
                return entities;
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<T>> GetWithFilterWithoutPartitionAsync(string filter)
        {
            try
            {
                var entities = await this.ExecuteQueryAsync(filter);
                return entities;
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<T>> GetAllAsync(string partition = null, int? count = null)
        {
            try
            {
                var partitionKeyFilter = this.GetPartitionKeyFilter(partition);
                var entities = await this.ExecuteQueryAsync(partitionKeyFilter, count);
                return entities;
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task<(IEnumerable<T>, string)> GetPagedAsync(string partition = null, int? count = null, string token = null)
        {
            var partitionKeyFilter = this.GetPartitionKeyFilter(partition);
            var pageable = this.Table.QueryAsync<T>(partitionKeyFilter, maxPerPage: count).AsPages(token);
            await using var enumerator = pageable.GetAsyncEnumerator();
            if (await enumerator.MoveNextAsync())
            {
                var page = enumerator.Current;
                return (page.Values, page.ContinuationToken);
            }

            return (Enumerable.Empty<T>(), null);
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<T>> GetAllLessThanDateTimeAsync(DateTime dateTime)
        {
            var filterByDate = $"Timestamp le datetime'{dateTime:o}'";
            var entities = await this.ExecuteQueryAsync(filterByDate);
            return entities;
        }

        /// <inheritdoc/>
        public async IAsyncEnumerable<IEnumerable<T>> GetStreamsAsync(string partition = null, int? count = null)
        {
            var partitionKeyFilter = this.GetPartitionKeyFilter(partition);
            await foreach (var page in this.Table.QueryAsync<T>(partitionKeyFilter, maxPerPage: count).AsPages())
            {
                yield return page.Values;
            }
        }

        /// <inheritdoc/>
        public async Task BatchInsertOrMergeAsync(IEnumerable<T> entities)
        {
            try
            {
                var array = entities.ToArray();
                for (var i = 0; i <= array.Length / 100; i++)
                {
                    var lowerBound = i * 100;
                    var upperBound = Math.Min(lowerBound + 99, array.Length - 1);
                    if (lowerBound > upperBound)
                    {
                        break;
                    }

                    var actions = new List<TableTransactionAction>();
                    for (var j = lowerBound; j <= upperBound; j++)
                    {
                        actions.Add(new TableTransactionAction(TableTransactionActionType.UpsertMerge, array[j]));
                    }

                    await this.Table.SubmitTransactionAsync(actions);
                }
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task BatchDeleteAsync(IEnumerable<T> entities)
        {
            var array = entities.ToArray();
            for (var i = 0; i <= array.Length / 100; i++)
            {
                var lowerBound = i * 100;
                var upperBound = Math.Min(lowerBound + 99, array.Length - 1);
                if (lowerBound > upperBound)
                {
                    break;
                }

                var actions = new List<TableTransactionAction>();
                for (var j = lowerBound; j <= upperBound; j++)
                {
                    actions.Add(new TableTransactionAction(TableTransactionActionType.Delete, array[j], ETag.All));
                }

                await this.Table.SubmitTransactionAsync(actions);
            }
        }

        /// <summary>
        /// Get a filter that filters in the entities matching the incoming row keys.
        /// </summary>
        /// <param name="rowKeys">Row keys.</param>
        /// <returns>A filter that filters in the entities matching the incoming row keys.</returns>
        protected string GetRowKeysFilter(IEnumerable<string> rowKeys)
        {
            try
            {
                var rowKeysFilter = string.Empty;
                foreach (var rowKey in rowKeys)
                {
                    var singleRowKeyFilter = $"RowKey eq '{rowKey}'";

                    if (string.IsNullOrWhiteSpace(rowKeysFilter))
                    {
                        rowKeysFilter = singleRowKeyFilter;
                    }
                    else
                    {
                        rowKeysFilter = $"({rowKeysFilter}) or ({singleRowKeyFilter})";
                    }
                }

                return rowKeysFilter;
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        private string CombineFilters(string filter1, string filter2)
        {
            if (string.IsNullOrWhiteSpace(filter1) && string.IsNullOrWhiteSpace(filter2))
            {
                return string.Empty;
            }
            else if (string.IsNullOrWhiteSpace(filter1))
            {
                return filter2;
            }
            else if (string.IsNullOrWhiteSpace(filter2))
            {
                return filter1;
            }

            return $"({filter1}) and ({filter2})";
        }

        private string GetPartitionKeyFilter(string partition)
        {
            var key = string.IsNullOrWhiteSpace(partition) ? this.defaultPartitionKey : partition;
            return $"PartitionKey eq '{key}'";
        }

        private async Task<IList<T>> ExecuteQueryAsync(string filter, int? count = null)
        {
            try
            {
                var result = new List<T>();
                await foreach (var entity in this.Table.QueryAsync<T>(filter, maxPerPage: count))
                {
                    result.Add(entity);
                    if (count.HasValue && result.Count >= count.Value)
                    {
                        break;
                    }
                }

                return result;
            }
            catch (RequestFailedException e)
            {
                this.Logger.LogError(e, e.Message);
                throw;
            }
        }
    }
}