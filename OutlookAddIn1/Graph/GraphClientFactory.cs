using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAddIn1.Graph
{
    internal sealed class MsalAccessTokenProvider : IAccessTokenProvider
    {
        private readonly IPublicClientApplication _pca;
        private readonly string[] _scopes;
        public AllowedHostsValidator AllowedHostsValidator { get; } = new AllowedHostsValidator();

        public MsalAccessTokenProvider(IPublicClientApplication pca, string[] scopes)
        {
            _pca = pca;
            _scopes = scopes;
        }

        public async Task<string> GetAuthorizationTokenAsync(
            Uri uri,
            Dictionary<string, object> additionalAuthenticationContext = null,
            CancellationToken cancellationToken = default)
        {
            var account = (await _pca.GetAccountsAsync()).FirstOrDefault();
            try
            {
                var res = await _pca.AcquireTokenSilent(_scopes, account).ExecuteAsync(cancellationToken);
                return res.AccessToken;
            }
            catch (MsalUiRequiredException)
            {
                var res = await _pca.AcquireTokenInteractive(_scopes).ExecuteAsync(cancellationToken);
                return res.AccessToken;
            }
        }
    }

    internal static class GraphClientFactory
    {
        public static async Task<GraphServiceClient> CreateAsync()
        {
            var pca = MsalClientFactory.Get();
            var tokenProvider = new MsalAccessTokenProvider(pca, AuthConfig.Scopes);
            var authProvider = new BaseBearerTokenAuthenticationProvider(tokenProvider);
            var client = new GraphServiceClient(authProvider);
            return await Task.FromResult(client);
        }
    }
}
