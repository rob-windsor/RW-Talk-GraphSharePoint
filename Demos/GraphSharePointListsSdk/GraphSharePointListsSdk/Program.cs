using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;
using System.Security.Cryptography.X509Certificates;

namespace GraphSharePointListsSdk
{
    internal class Program
    {
        private static string siteName = "M365NYCPrep";

        static void Main(string[] args)
        {
            GetLists().Wait();
        }

        private async static Task GetLists()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteUrl = "robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:";

                var result = await graphClient.Sites[siteUrl]
                    .Lists
                    .GetAsync();

                foreach (var list in result.Value)
                {
                    Console.WriteLine(list.DisplayName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private async static Task GetList()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteUrl = "robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:";

                var list = await graphClient.Sites[siteUrl]
                    .Lists["Products"]
                    .GetAsync((config) =>
                    {
                        config.QueryParameters.Expand = 
                            new string[] { "columns" };
                    });

                Console.WriteLine(list.DisplayName);
                foreach (var column in list.Columns)
                {
                    if (column.Hidden == true || column.ReadOnly == true)
                    {
                        continue;
                    }

                    Console.WriteLine("\t" + column.DisplayName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private async static Task CreateList()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteUrl = "robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:";

                var body = new Microsoft.Graph.Models.List
                {
                    DisplayName = "Products" + DateTime.Now.Ticks,
                    Columns = new List<ColumnDefinition>
                    {
                        new ColumnDefinition
                        {
                            Name = "Category",
                            Lookup = new LookupColumn
                            {
                                ColumnName = "Title",
                                ListId = "c1a273f7-e422-4f2f-8310-f6d6370b8410"
                            },
                        },
                        new ColumnDefinition
                        {
                            Name = "QuantityPerUnit",
                            Text = new TextColumn
                            {
                            },
                        },
                        new ColumnDefinition
                        {
                            Name = "UnitPrice",
                            Currency = new CurrencyColumn
                            {
                                Locale = "en-US"
                            },
                        },
                        new ColumnDefinition
                        {
                            Name = "UnitsInStock",
                            Number = new NumberColumn
                            {
                            },
                        },
                        new ColumnDefinition
                        {
                            Name = "Discontinued",
                            Boolean = new BooleanColumn
                            {
                            },
                        }
                    },
                    ListProp = new ListInfo
                    {
                        Template = "genericList"
                    }
                };

                var list = await graphClient.Sites[siteUrl]
                    .Lists
                    .PostAsync(body);

                Console.WriteLine(list.DisplayName);
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
