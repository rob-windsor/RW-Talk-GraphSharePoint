using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json.Nodes;
using System.Text.Json;
using System.Net.Http.Json;

namespace GraphSharePointLists
{
    internal class Program
    {
        private static string siteName = "M365NYCPrep";
        private static HttpClient httpClient = new HttpClient();

        static void Main(string[] args)
        {
            GetLists().Wait();
        }

        private async static Task GetLists()
        {
            try
            {
                var token = await GetAccessToken();

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                var apiUrl = "https://graph.microsoft.com/v1.0/sites" +
                    "/robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:" +
                    "/lists";
                using (var response = await httpClient.GetAsync(apiUrl))
                {
                    response.EnsureSuccessStatusCode();

                    var jsonText = await response.Content.ReadAsStringAsync();
                    var jsonNode = JsonNode.Parse(jsonText);
                    var options = new JsonSerializerOptions() { WriteIndented = true };
                    Console.WriteLine(jsonNode!.ToJsonString(options));
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
                var token = await GetAccessToken();

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                var apiUrl = "https://graph.microsoft.com/v1.0/sites" +
                    "/robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:" +
                    "/lists/Products" +
                    "?expand=columns";
                using (var response = await httpClient.GetAsync(apiUrl))
                {
                    response.EnsureSuccessStatusCode();

                    var jsonText = await response.Content.ReadAsStringAsync();
                    var jsonNode = JsonNode.Parse(jsonText);
                    var options = new JsonSerializerOptions() { WriteIndented = true };
                    Console.WriteLine(jsonNode!.ToJsonString(options));
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
                var token = await GetAccessToken();

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                var apiUrl = "https://graph.microsoft.com/v1.0/sites" +
                    "/robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:" +
                    "/lists";
                var bodyObj = new
                {
                    displayName = "Products" + DateTime.Now.Ticks,
                    columns = new dynamic[] {
                        new
                        {
                            name = "Category",
                            lookup = new
                            {
                                columnName = "Title",
                                listId = "c1a273f7-e422-4f2f-8310-f6d6370b8410"
                            }
                        },
                        new
                        {
                            name = "QuantityPerUnit",
                            text = new { }
                        },
                        new
                        {
                            name = "UnitPrice",
                            currency = new
                            {
                                locale = "en-US"
                            }
                        },
                        new
                        {
                            name = "UnitsInStock",
                            number = new { }
                        },
                        new
                        {
                            name = "Discontinued",
                            boolean = new { }
                        }
                    },
                    list = new
                    {
                        template = "genericList"
                    }
                };
                var body = JsonContent.Create(bodyObj);
                using (var response = await httpClient.PostAsync(apiUrl, body))
                {
                    response.EnsureSuccessStatusCode();

                    var jsonText = await response.Content.ReadAsStringAsync();
                    var jsonNode = JsonNode.Parse(jsonText);
                    var options = new JsonSerializerOptions() { WriteIndented = true };
                    Console.WriteLine(jsonNode!.ToJsonString(options));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
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
