// <copyright file="GraphHelper.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>

namespace Sample.IncidentBot.Helpers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Microsoft.Graph.Auth;
    using Microsoft.Identity.Client;
    using Sample.IncidentBot.Interface;

    /// <summary>
    /// Graph helper.
    /// </summary>
    public class GraphHelper : IGraph
    {
        /// <summary>
        /// Azure Client Id.
        /// </summary>
        public const string AppIdConfigurationSettingsKey = "AzureAd:AppId";

        /// <summary>
        /// Azure Tenant Id.
        /// </summary>
        public const string TenantIdConfigurationSettingsKey = "AzureAd:TenantId";

        /// <summary>
        /// Azure ClientSecret .
        /// </summary>
        public const string AppSecretConfigurationSettingsKey = "AzureAd:AppSecret";

        private readonly ILogger<GraphHelper> logger;
        private readonly IConfiguration configuration;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphHelper"/> class.
        /// </summary>
        /// <param name="logger">ILogger instance.</param>
        /// <param name="configuration">IConfiguration instance.</param>
        public GraphHelper(ILogger<GraphHelper> logger, IConfiguration configuration)
        {
            this.logger = logger;
            this.configuration = configuration;
        }

        /// <inheritdoc/>
        public async Task<OnlineMeeting> CreateOnlineMeetingAsync(GraphServiceClient graphServiceClient, OnlineMeeting onlineMeeting)
        {
            try
            {
                var onlineMeetingResponse = await graphServiceClient.Users[this.configuration["UserId"]].OnlineMeetings
                    .Request()
                    .AddAsync(onlineMeeting).ConfigureAwait(false);
                return onlineMeetingResponse;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Get graph service client information.
        /// </summary>
        /// <returns>Graph service client.</returns>
        public GraphServiceClient GetGraphServiceClient()
        {
            try
            {
                return new GraphServiceClient(GetClientCredentialProvider());
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Get Client Credential Provider.
        /// </summary>
        /// <returns>Client Credential Provider.</returns>
        private static ClientCredentialProvider GetClientCredentialProvider()
        {
            try
            {
                IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                        .Create(Startup.StaticConfiguration[AppIdConfigurationSettingsKey])
                        .WithTenantId(Startup.StaticConfiguration[TenantIdConfigurationSettingsKey])
                        .WithClientSecret(Startup.StaticConfiguration[AppSecretConfigurationSettingsKey])
                        .Build();

                return new ClientCredentialProvider(confidentialClientApplication);
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }
    }
}
