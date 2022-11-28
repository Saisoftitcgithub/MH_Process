using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using System.Text;
using Azure.Identity;
using System.IO.Pipes;
using Microsoft.Extensions.Azure;
using static System.Net.WebRequestMethods;
using daemon_console;

namespace One_Drive_Process
{
    class ExternalAPIMethods
    {
        public void UploadOneDriveFile(string filePath)
        {
            try
            {
                RunAsync(filePath).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }

        private static async Task RunAsync(string filePath)
        {
            AuthenticationConfig config = AuthenticationConfig.ReadFromJsonFile("appsettings.json");

            // You can run this sample using ClientSecret or Certificate. The code will differ only when instantiating the IConfidentialClientApplication
            bool isUsingClientSecret = IsAppUsingClientSecret(config);

            // Even if this is a console application here, a daemon application is a confidential client application
            IConfidentialClientApplication app;

            if (isUsingClientSecret)
            {
                // Even if this is a console application here, a daemon application is a confidential client application
                app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
                    .WithClientSecret(config.ClientSecret)
                    .WithAuthority(new Uri(config.Authority))
                    .Build();
            }

            else
            {
                ICertificateLoader certificateLoader = new DefaultCertificateLoader();
                certificateLoader.LoadIfNeeded(config.Certificate);

                app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
                    .WithCertificate(config.Certificate.Certificate)
                    .WithAuthority(new Uri(config.Authority))
                    .Build();
            }

            app.AddInMemoryTokenCache();

            // With client credentials flows the scopes is ALWAYS of the shape "resource/.default", as the 
            // application permissions need to be set statically (in the portal or by PowerShell), and then granted by
            // a tenant administrator. 
            string[] scopes = new string[] { $"{config.ApiUrl}.default" }; // Generates a scope -> "https://graph.microsoft.com/.default"
            //List<string> scopes = new List<string>();
            //scopes.Add("User.Read");
            //scopes.Add("Files.Read");
            //scopes.Add("Files.ReadWrite");
            // Call MS graph using the Graph SDK
            await CallMSGraphUsingGraphSDK(app, scopes, filePath);

            // Call MS Graph REST API directly
            await CallMSGraph(config, app, scopes);
        }

        private static async Task CallMSGraph(AuthenticationConfig config, IConfidentialClientApplication app, string[] scopes)
        {
            AuthenticationResult result = null;
            try
            {
                result = await app.AcquireTokenForClient(scopes)
                    .ExecuteAsync();

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Token acquired");
                Console.ResetColor();
            }
            catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS70011"))
            {
                // Invalid scope. The scope has to be of the form "https://resourceurl/.default"
                // Mitigation: change the scope to be as expected
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Scope provided is not supported");
                Console.ResetColor();
            }

            // The following example uses a Raw Http call 
            if (result != null)
            {
                var httpClient = new HttpClient();
                var apiCaller = new ProtectedApiCallHelper(httpClient);
                await apiCaller.CallWebApiAndProcessResultASync($"{config.ApiUrl}v1.0/users", result.AccessToken, Display);

            }
        }
        private static async Task CallMSGraphUsingGraphSDK(IConfidentialClientApplication app, string[] scopes, string filePath)
        {
            // Prepare an authenticated MS Graph SDK client
            GraphServiceClient graphServiceClient = GetAuthenticatedGraphClient(app, scopes);


            List<User> allUsers = new List<User>();

            try
            {

                IGraphServiceUsersCollectionPage users = await graphServiceClient.Users.Request().GetAsync();
                Console.WriteLine($"Found {users.Count()} users in the tenant");
            }
            catch (ServiceException e)
            {
                Console.WriteLine("We could not retrieve the user's list: " + $"{e}");
            }

            //DateTime localDate = DateTime.Now;
            var CurrentDateTime = DateTime.Now.ToString("yyyyMMddHHmmss").Replace('/', ' ');
            //var CurrentDate = DateTime.Now.Date.ToString("yyyy.MM.dd").Replace('/', ' ');
            var folderName = "MH DOCS " + CurrentDateTime;

            //create folder
            try
            {


                var driveItem = new DriveItem
                {

                    Name = folderName,
                    Folder = new Folder
                    {
                    },
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"@microsoft.graph.conflictBehavior", "rename"}
                    }
                };

                var userId = "bd77c85b-c847-4b38-85de-3fce6ba46332";
                await graphServiceClient.Users[userId].Drive.Root.Children
                                        .Request()
                                        .AddAsync(driveItem);

                /* var folder = await graphServiceClient.Users[userId]
                                             .Drive.Root
                                             .ItemWithPath("/New Folder")
                                             .Request()
                                             .GetAsync();
                 var itemId = folder.Id;
                 Console.WriteLine(folder.Id);*/
            }
            catch (Exception e)
            {
                Console.WriteLine("Folder creation exception: " + e);
            }
            //Upload the file
            try
            {

                //var filename = @"C:\Office\Soujanya\VisualStudio\MH_TAXINVOICE_PROCESS\MH_INVOICE_V1\wi-476493_Compressed.pdf";
                var filename = "wi-476493_Compressed.pdf";
                //var filename = @"\testupload.txt";
                //var filePath = Path.Combine(System.IO.Directory.GetCurrentDirectory(), filename);
                Console.WriteLine("Uploading file: " + filename);
                var userId = "bd77c85b-c847-4b38-85de-3fce6ba46332";

                var folder = await graphServiceClient.Users[userId]
                                            .Drive.Root
                                            .ItemWithPath(folderName)
                                            .Request()
                                            .GetAsync();
                var itemId = folder.Id;
                Console.WriteLine(folder.Id);

                var UploadedFile = await graphServiceClient.Users[userId].Drive.Items[itemId]
                                   .ItemWithPath(filename)
                                   .CreateUploadSession()
                                   .Request()
                                   .PostAsync();

                using var stream = System.IO.File.OpenRead(filename);

                //create upload task
                var maxChunkSize = 320 * 1024;
                var largeUploadTask = new LargeFileUploadTask<DriveItem>(UploadedFile, stream, maxChunkSize);

                //create progress implementation
                IProgress<long> uploadProgress = new Progress<long>(uploadBytes =>
                {
                    try
                    {
                        Console.WriteLine($"Uploaded {uploadBytes} bytes of {stream.Length} bytes");
                    }
                    catch (Exception exce)
                    {
                        Console.WriteLine(exce);
                    }
                });

                //upload file
                try
                {
                    UploadResult<DriveItem> uploadResult = await largeUploadTask.UploadAsync(uploadProgress);
                    if (uploadResult.UploadSucceeded)
                    {
                        Console.WriteLine("file uploaded to one drive root folder");
                    }
                    else
                    {
                        Console.WriteLine("Upload Failed");
                    }

                }

                catch (ServiceException ex)
                {
                    Console.WriteLine($"Error uploading: {ex.ToString()}");
                }

            }



            catch (Exception exe)
            {
                Console.WriteLine(exe);
            }

        }





        /// <summary>
        /// An example of how to authenticate the Microsoft Graph SDK using the MSAL library
        /// </summary>
        /// <returns></returns>
        private static GraphServiceClient GetAuthenticatedGraphClient(IConfidentialClientApplication app, string[] scopes)
        {

            GraphServiceClient graphServiceClient =
                    new GraphServiceClient("https://graph.microsoft.com/V1.0/", new DelegateAuthenticationProvider(async (requestMessage) =>
                    {
                        // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                        AuthenticationResult result = await app.AcquireTokenForClient(scopes)
                            .ExecuteAsync();

                        // Add the access token in the Authorization header of the API request.
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    }));

            return graphServiceClient;
        }

        private static void Display(JsonNode result)
        {
            JsonArray nodes = ((result as JsonObject).ToArray()[1]).Value as JsonArray;

            foreach (JsonObject aNode in nodes.ToArray())
            {
                foreach (var property in aNode.ToArray())
                {
                    Console.WriteLine($"{property.Key} = {property.Value?.ToString()}");
                }
                Console.WriteLine();
            }
        }

        private static bool IsAppUsingClientSecret(AuthenticationConfig config)
        {
            string clientSecretPlaceholderValue = "[Enter here a client secret for your application]";

            if (!String.IsNullOrWhiteSpace(config.ClientSecret) && config.ClientSecret != clientSecretPlaceholderValue)
            {
                return true;
            }

            else if (config.Certificate != null)
            {
                return false;
            }

            else
                throw new Exception("You must choose between using client secret or certificate. Please update appsettings.json file.");
        }

    }
}
