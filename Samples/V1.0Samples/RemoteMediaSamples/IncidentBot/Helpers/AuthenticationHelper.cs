// <copyright file="AuthenticationHelper.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
// </copyright>

namespace Sample.IncidentBot.Helpers
{
    using Microsoft.AspNetCore.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Graph;
    using Microsoft.Graph.Auth;
    using Microsoft.Identity.Client;

    /// <summary>
    /// Authentication helper.
    /// </summary>
    public static class AuthenticationHelper
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

        /// <summary>
        /// GraphServiceClient.
        /// </summary>
        /// <returns>Graph Service Client.</returns>
        public static GraphServiceClient GetGraphServiceClient()
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
