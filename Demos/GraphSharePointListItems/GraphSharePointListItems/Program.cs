using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json.Nodes;
using System.Text.Json;
using System.Net.Http.Json;
using System.Text;

namespace GraphSharePointListItems
{
    internal class Program
    {
        private static string siteName = "M365NYCPrep";

        private static HttpClient httpClient = new HttpClient();

        static void Main(string[] args)
        {
            GetListItems().Wait();
        }

        private async static Task GetListItems()
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
                    "/items" +
                    "?expand=fields";
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

        private async static Task GetListItemsWithFilter()
        {
            try
            {
                var token = await GetAccessToken();

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
                httpClient.DefaultRequestHeaders.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");

                var apiUrl = "https://graph.microsoft.com/v1.0/sites" +
                    "/robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:" +
                    "/lists/Products" +
                    "/items" +
                    "?expand=fields" +
                    "&filter=fields/CategoryLookupId eq 1 and fields/Discontinued eq false";
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

        private async static Task GetListItem()
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
                    "/items/1";
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

        private async static Task GetListItemWithSelectAndExpand()
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
                    "/items/1" +
                    "?expand=fields(select=Id,Title,Category,QuantityPerUnit,UnitPrice,UnitsInStock)";
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

        private async static Task CreateListItem()
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
                    "/lists/Categories" +
                    "/items";
                var bodyStr =
                    "{" +
                    "    'fields': {" +
                    $"       'Title': 'Category {DateTime.Now.Ticks}'," +
                    "        'Description': 'This is a test'" +
                    "    }" +
                    "}";
                var body = new StringContent(bodyStr, Encoding.UTF8, "application/json");
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

        private async static Task CreateListItemWithLookups()
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
                    "/lists/LookupTest" +
                    "/items";
                var bodyStr =
                    "{" +
                    "    'fields': {" +
                    $"        'Title': 'Item {DateTime.Now.Ticks}'," +
                    "        'ColorLookupId': 1," +
                    "        'ColorsLookupId@odata.type': 'Collection(Edm.Int32)'," +
                    "        'ColorsLookupId': [1, 3]" +
                    "    }" +
                    "}";
                var body = new StringContent(bodyStr, Encoding.UTF8, "application/json");
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
