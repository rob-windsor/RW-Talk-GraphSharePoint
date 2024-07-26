using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;
using System.Security.Cryptography.X509Certificates;

namespace GraphSharePointSitesSdk
{
    internal class Program
    {
        private static string siteName = "M365NYCPrep";

        static void Main(string[] args)
        {
            SearchSites().Wait();
        }

        private async static Task SearchSites()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var result = await graphClient.Sites
                    .GetAsync((config) =>
                    {
                        config.QueryParameters.Search = "\"M365*\"";
                    });

                foreach (var site in result.Value)
                {
                    Console.WriteLine(site.DisplayName);
                    Console.WriteLine(site.Id);
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private async static Task GetSiteByUrl()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteUrl = "robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}";

                var site = await graphClient.Sites[siteUrl]
                    .GetAsync();

                Console.WriteLine(site.DisplayName);
                Console.WriteLine(site.Id);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private async static Task GetSiteById()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteId = "robwindsortest980.sharepoint.com," +
                    "b9f16cf4-48ae-44e0-abc7-b4ecbd68b879," +
                    "256de9de-e9de-43c2-aa6f-9bd9f4d9f619";

                var site = await graphClient.Sites[siteId]
                    .GetAsync();

                Console.WriteLine(site.DisplayName);
                Console.WriteLine(site.Id);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }


        private class TokenProvider : IAccessTokenProvider
        {
            public async Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object> additionalAuthenticationContext = default,
                CancellationToken cancellationToken = default)
            {
                var token = await GetAccessToken();
                return token;
            }

            public AllowedHostsValidator AllowedHostsValidator { get; }
        }

        private async static Task<string> GetAccessToken()
        {
            var certPath = @"C:\Dev\Community\Presentations\RW-Talk-GraphSharePoint\Demos\Certificate"; ;
            var certName = "M365NYC.pfx";
            var certPassword = "pass@word1";
            var cert = new X509Certificate2(
                System.IO.File.ReadAllBytes(Path.Combine(certPath, certName)),
                certPassword,
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet);

            var clientId = "352ed383-ac24-4dfe-9cd9-5ccc66f2eaed";
            var authority = "https://login.microsoftonline.com/robwindsortest980.onmicrosoft.com/";
            var azureApp = ConfidentialClientApplicationBuilder.Create(clientId)
                .WithAuthority(authority)
                .WithCertificate(cert)
                .Build();

            var scopes = new string[] { $"https://graph.microsoft.com/.default" };
            var authResult = await azureApp.AcquireTokenForClient(scopes).ExecuteAsync();
            return authResult.AccessToken;
        }
    }
}
