/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Security.Claims;

namespace MicrosoftGraphAspNetCoreConnectSample.Services
{
    public class GraphServiceClientFactory : IGraphServiceClientFactory
    {
        private readonly IGraphAuthProvider _authProvider;

        public GraphServiceClientFactory(IGraphAuthProvider authProvider)
        {
            _authProvider = authProvider;
        }

        public GraphServiceClient GetAuthenticatedGraphClient(ClaimsIdentity userIdentity) =>
            new GraphServiceClient(new DelegateAuthenticationProvider(
                async requestMessage =>
                {
                    // Get user's id for token cache.
                    var identifier = userIdentity.FindFirst(Startup.ObjectIdentifierType)?.Value + "." + userIdentity.FindFirst(Startup.TenantIdType)?.Value;

                    // Passing tenant ID to the sample auth provider to use as a cache key
                    var accessToken = await _authProvider.GetUserAccessTokenAsync(identifier);

                    // Append the access token to the request
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    // This header identifies the sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    requestMessage.Headers.Add("SampleID", "aspnetcore-connect-sample");
                }));
    }

    public interface IGraphServiceClientFactory
    {
        GraphServiceClient GetAuthenticatedGraphClient(ClaimsIdentity userIdentity);
    }
}
