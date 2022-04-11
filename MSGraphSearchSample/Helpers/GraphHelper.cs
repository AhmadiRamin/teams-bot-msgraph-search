using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using MSGraphSearchSample.Interfaces;
using MSGraphSearchSample.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace MSGraphSearchSample.Helpers
{
    public class GraphHelper : IGraphHelper
    {
        protected readonly AppConfigOptions _appconfig;
        public GraphHelper(IOptions<AppConfigOptions> options)
        {
            _appconfig = options.Value;
        }
        public GraphServiceClient GetApplicationServiceClient()
        {
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
            .Create(_appconfig.MicrosoftAppId)
            .WithTenantId(_appconfig.MicrosoftAppTenantId)
            .WithClientSecret(_appconfig.MicrosoftAppPassword)
            .Build();
            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);
            return graphClient;
        }

        public GraphServiceClient GetDelegatedServiceClient(string _token)
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    requestMessage =>
                    {
                        // Append the access token to the request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", _token);

                        // Get event times in the current time zone.
                        requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                        return Task.CompletedTask;
                    }));
            return graphClient;
        }

        public async Task<string> GetOnBehalfOfAccessToken(string _token, string resourceUri)
        {
            List<string> scopes = new List<string>();
            scopes.Add(resourceUri);
            var app = ConfidentialClientApplicationBuilder.Create(_appconfig.MicrosoftAppId)
                .WithClientSecret(_appconfig.MicrosoftAppPassword)
                .WithTenantId(_appconfig.MicrosoftAppTenantId)
                .Build();
            var userAssertion = new UserAssertion(_token);
            var result = await app.AcquireTokenOnBehalfOf(scopes, userAssertion).ExecuteAsync();
            return result.AccessToken;
        }
    }
}
