// <copyright file="IGroupMembersService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Graph.Models;

    /// <summary>
    /// Interface for Group Members Service.
    /// </summary>
    public interface IGroupMembersService
    {
        /// <summary>
        /// Get groups members.
        /// </summary>
        /// <param name="groupId">Group Id.</param>
        /// <returns>Enumerator to iterate over a collection of <see cref="User"/>.</returns>
        Task<IEnumerable<User>> GetGroupMembersAsync(string groupId);
    }
}
