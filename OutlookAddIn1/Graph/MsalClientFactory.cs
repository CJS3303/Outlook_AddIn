using Microsoft.Identity.Client;

namespace OutlookAddIn1.Graph
{
    internal static class MsalClientFactory
    {
        private static IPublicClientApplication _pca;

        public static IPublicClientApplication Get()
        {
            if (_pca != null) return _pca;

            _pca = PublicClientApplicationBuilder
                .Create(AuthConfig.ClientId)
                .WithTenantId(AuthConfig.TenantId)
                .WithRedirectUri(AuthConfig.RedirectUri)
                // Optional: use Windows Account Manager SSO on corp machines:
                //.WithWindowsBroker(true) 
                .Build();

            return _pca;
        }
    }
}
