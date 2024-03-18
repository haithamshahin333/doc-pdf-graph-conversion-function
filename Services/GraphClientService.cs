// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace Company.Services
{
    public class GraphClientService : IGraphClientService
    {
        private readonly IConfiguration _config;
        private readonly ILogger _logger;
        private GraphServiceClient? _appGraphClient;

        public GraphClientService(IConfiguration config, ILoggerFactory loggerFactory)
        {
            _config = config;
            _logger = loggerFactory.CreateLogger<GraphClientService>();
        }

        public GraphServiceClient? GetAppGraphClient()
        {

            _logger.LogInformation("Obtaining Graph Client");

            string userAssignedClientId = _config.GetValue<String>("CLIENT_APP_ID");

            if (_appGraphClient == null)
            {

                if (string.IsNullOrEmpty(userAssignedClientId))
                {
                    _logger.LogError("Required settings missing: 'CLIENT_APP_ID'");
                    return null;
                }

                var credential = new DefaultAzureCredential(new DefaultAzureCredentialOptions { ManagedIdentityClientId = userAssignedClientId });
                _appGraphClient = new GraphServiceClient(credential);
            }

            return _appGraphClient;
        }
    }
}