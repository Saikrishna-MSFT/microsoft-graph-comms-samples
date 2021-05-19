// <copyright file="IGraph.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>

namespace Sample.IncidentBot.Interface
{
    using System.Threading.Tasks;
    using Microsoft.Graph;

    /// <summary>
    /// Interface for Graph.
    /// </summary>
    public interface IGraph
    {
        /// <summary>
        /// Get graph service client.
        /// </summary>
        /// <returns>Graph service client.</returns>
        GraphServiceClient GetGraphServiceClient();

        /// <summary>
        /// Creates online meeting.
        /// </summary>
        /// <param name="graphServiceClient">GraphServiceClient instance.</param>
        /// <param name="onlineMeeting">OnlineMeeting instance.</param>
        /// <returns>Online meeting details.</returns>
        Task<OnlineMeeting> CreateOnlineMeetingAsync(GraphServiceClient graphServiceClient, OnlineMeeting onlineMeeting);
    }
}
