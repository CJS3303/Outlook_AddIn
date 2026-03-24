using Microsoft.IdentityModel.Protocols;
using System.Configuration;


namespace OutlookAddIn1.Graph
{
    internal static class AuthConfig
    {
        public static string ClientId => ConfigurationManager.AppSettings["ClientId"];
        public static string TenantId => ConfigurationManager.AppSettings["TenantId"];
        public static string RedirectUri => ConfigurationManager.AppSettings["RedirectUri"];
        public static string[] Scopes => new[] {
            "User.Read", "Calendars.ReadWrite", "OnlineMeetings.ReadWrite"
        };
    }
}
