using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json.Nodes;
using System.Text.Json;

namespace GraphSharePointSites
{
    internal class Program
    {
        private static string siteName = "M365NYCPrep";
        private static HttpClient httpClient = new HttpClient();

        static void Main(string[] args)
        {
            SearchSites().Wait();
        }

        private async static Task SearchSites()
        {
            try
            {
                var token = await GetAccessToken();

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                var apiUrl = "https://graph.microsoft.com/v1.0/sites" +
                    "?search='M365*'";
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

        private async static Task GetSiteByUrl()
        {
            try
            {
                var token = await GetAccessToken();

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                var apiUrl = "https://graph.microsoft.com/v1.0/sites" +
                    "/robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}";
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

        private async static Task GetSiteById()
        {
            try
            {
                var token = await GetAccessToken();

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                var apiUrl = "https://graph.microsoft.com/v1.0/sites" +
                    "/robwindsortest980.sharepoint.com," +
                    "b9f16cf4-48ae-44e0-abc7-b4ecbd68b879," +
                    "256de9de-e9de-43c2-aa6f-9bd9f4d9f619";
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
