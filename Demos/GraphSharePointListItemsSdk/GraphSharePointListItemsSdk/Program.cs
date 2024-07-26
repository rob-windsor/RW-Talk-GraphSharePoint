using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;
using System.Security.Cryptography.X509Certificates;

namespace GraphSharePointListItemsSdk
{
    internal class Program
    {
        private static string siteName = "M365NYCPrep";

        static void Main(string[] args)
        {
            GetListItems().Wait();
        }

        private async static Task GetListItems()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteUrl = "robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:";

                var result = await graphClient.Sites[siteUrl]
                    .Lists["Products"]
                    .Items
                    .GetAsync();

                foreach (var item in result.Value)
                {
                    Console.WriteLine(item.Id);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private async static Task GetListItemsWithFilter()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteUrl = "robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:";

                var result = await graphClient.Sites[siteUrl]
                    .Lists["Products"]
                    .Items
                    .GetAsync((config) =>
                    {
                        config.Headers.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");

                        config.QueryParameters.Expand = new string[] { "fields" };
                        config.QueryParameters.Filter =
                            "fields/CategoryLookupId eq 1 and fields/Discontinued eq false";
                    });

                foreach (var item in result.Value)
                {
                    Console.WriteLine(item.Id);
                    foreach (var fieldData in item.Fields.AdditionalData)
                    {
                        Console.WriteLine($"\t{fieldData.Key}={fieldData.Value}");
                    }
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private async static Task GetListItem()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteUrl = "robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:";

                var item = await graphClient.Sites[siteUrl]
                    .Lists["Products"]
                    .Items["1"]
                    .GetAsync();

                Console.WriteLine(item.Id);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private async static Task GetListItemWithSelectAndExpand()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteUrl = "robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:";

                var item = await graphClient.Sites[siteUrl]
                    .Lists["Products"]
                    .Items["1"]
                    .GetAsync((config) =>
                    {
                        config.QueryParameters.Expand = new string[] 
                        {
                            "fields($select=Id,Title,CategoryLookupId,Category,QuantityPerUnit,UnitPrice,UnitsInStock)"
                        };
                    });

                Console.WriteLine(item.Id);
                foreach (var fieldData in item.Fields.AdditionalData)
                {
                    Console.WriteLine($"\t{fieldData.Key}={fieldData.Value}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private async static Task CreateListItem()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteUrl = "robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:";

                var body = new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>
                        {
                            {
                                "Title",
                                "Category " + DateTime.Now.Ticks
                            },
                            {
                                "Description",
                                "This is a test"
                            }
                        }
                    }
                };

                var item = await graphClient.Sites[siteUrl]
                    .Lists["Categories"]
                    .Items
                    .PostAsync(body);

                Console.WriteLine(item.Id);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private async static Task CreateListItemWithLookups()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteUrl = "robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:";

                var body = new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>
                        {
                            {
                                "Title",
                                "Item " + DateTime.Now.Ticks
                            },
                            {
                                "ColorLookupId",
                                1
                            },
                            {
                                "ColorsLookupId@odata.type",
                                "Collection(Edm.String)"
                            },
                            {
                                "ColorsLookupId",
                                new []{"1", "3"}
                            }
                        }
                    }
                };

                var item = await graphClient.Sites[siteUrl]
                    .Lists["LookupTest"]
                    .Items
                    .PostAsync(body);

                Console.WriteLine(item.Id);
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
