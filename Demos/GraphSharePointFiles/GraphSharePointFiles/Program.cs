using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json.Nodes;
using System.Text.Json;

namespace GraphSharePointFiles
{
    internal class Program
    {
        private static string siteName = "M365NYCPrep";
        private static string driveId = "b!9Gzxua5I4ESrx7TsvWi4ed7pbSXe6cJDqm-b2fTZ9hnVJ3SPa7Y7R5M9YqUIMAri";

        private static HttpClient httpClient = new HttpClient();

        static void Main(string[] args)
        {
            GetDocumentLibrary().Wait();
        }

        private async static Task GetDocumentLibraries()
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
                    "/drives";
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

        private async static Task GetDocumentLibrary()
        {
            try
            {
                var token = await GetAccessToken();

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                var apiUrl = "https://graph.microsoft.com/v1.0/drives" +
                    $"/{driveId}";
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

        private async static Task GetFiles()
        {
            try
            {
                var token = await GetAccessToken();

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                var apiUrl = "https://graph.microsoft.com/v1.0/drives" +
                    $"/{driveId}" +
                    "/root" +
                    "/children";
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

        private async static Task GetFile()
        {
            try
            {
                var token = await GetAccessToken();

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                var apiUrl = "https://graph.microsoft.com/v1.0/drives" +
                    $"/{driveId}" +
                    "/root:" +
                    "/SampleDocument001.docx";
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

        private async static Task UploadFile()
        {
            try
            {
                var filePath = @"C:\Users\Rob Windsor\Documents\Sample Documents";
                var fileName = "RobTest001.txt";
                var fileContents = File.ReadAllText(Path.Combine(filePath, fileName));

                var token = await GetAccessToken();

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                var apiUrl = "https://graph.microsoft.com/v1.0/drives" +
                    $"/{driveId}" +
                    "/root:" +
                    $"/RobTest{DateTime.Now.Ticks}.txt:" +
                    "/content";
                var body = new StringContent(fileContents);
                using (var response = await httpClient.PutAsync(apiUrl, body))
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
