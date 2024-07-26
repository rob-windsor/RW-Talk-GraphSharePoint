using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace GraphSharePointFilesSdk
{
    internal class Program
    {
        private static string siteName = "M365NYCPrep";
        private static string driveId = "b!9Gzxua5I4ESrx7TsvWi4ed7pbSXe6cJDqm-b2fTZ9hnVJ3SPa7Y7R5M9YqUIMAri";

        static void Main(string[] args)
        {
            GetDocumentLibraries().Wait();
        }

        private async static Task GetDocumentLibraries()
        {
            try
            {
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var siteUrl = "robwindsortest980.sharepoint.com:" +
                    $"/sites/{siteName}:";

                var result = await graphClient.Sites[siteUrl]
                    .Drives
                    .GetAsync();

                foreach (var drive in result.Value)
                {
                    Console.WriteLine(drive.Name);
                    Console.WriteLine(drive.Id);
                    Console.WriteLine();
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
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var drive = await graphClient.Drives[driveId]
                    .GetAsync();

                Console.WriteLine(drive.Name);
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
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var result = await graphClient.Drives[driveId]
                    .Items["root"]
                    .Children
                    .GetAsync();

                foreach (var item in result.Value)
                {
                    Console.WriteLine(item.Name);
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
                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var item = await graphClient.Drives[driveId]
                    .Items["root"]
                    .ItemWithPath("SampleDocument001.docx")
                    .GetAsync();

                Console.WriteLine(item.Name);
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

                var authenticationProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider());
                var graphClient = new GraphServiceClient(authenticationProvider);

                var body = new MemoryStream(Encoding.UTF8.GetBytes(fileContents));
                var item = await graphClient.Drives[driveId]
                    .Items["root"]
                    .ItemWithPath($"RobTest{DateTime.Now.Ticks}.txt")
                    .Content
                    .PutAsync(body);

                Console.WriteLine(item.Name);
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
