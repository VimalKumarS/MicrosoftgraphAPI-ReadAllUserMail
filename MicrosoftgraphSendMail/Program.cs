using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft_Graph_Mail_Console_App;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftgraphSendMail
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //certificateAuthentication();
            var token = certificateAuthentication_graph();
            //GetMailMessage(token);
            creategraphclient(token);
            //MailClient.SendMeAsync().Wait();
            //Console.Read();
        }

        //get outlook credential
        private static void certificateAuthentication()
        {
            string tenantId = "tesla329.onmicrosoft.com";
            string clientId = "dd784c4b-2956-44ae-9974-d84caf0e3d96";
            string resourceId = "https://outlook.office.com/";
            string resourceUrl = "https://outlook.office.com/api/v2.0/users/user1@tesla329.onmicrosoft.com/sendmail"; //this is your on-behalf user's UPN
            string authority = String.Format("https://login.windows.net/{0}", tenantId);
            string certficatePath = @"C:\Dexter\Practice\outlook_mail\warpbubble.pfx"; //this is your certficate location.
            string certificatePassword = "p@ssw0rd"; // this is your certificate password
            X509Certificate2 certificate = new X509Certificate2(certficatePath, certificatePassword, X509KeyStorageFlags.MachineKeySet);
            AuthenticationContext authenticationContext = new AuthenticationContext(authority, false);

            ClientAssertionCertificate cac = new ClientAssertionCertificate(clientId, certificate);

            //get the access token to Outlook using the ClientAssertionCertificate
            var authenticationResult = authenticationContext.AcquireTokenAsync(resourceId, cac).Result;
            string token = authenticationResult.AccessToken;

            var itemPayload = new
            {
                Message = new
                {
                    Subject = "Test email",
                    Body = new { ContentType = "Text", Content = "this is test email." },
                    ToRecipients = new[] { new { EmailAddress = new { Address = "warpbubble@tesla329.onmicrosoft.com" } } }
                }
            };

            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
            client.DefaultRequestHeaders.Add("Accept", "application/json");
            HttpContent content = new StringContent(JsonConvert.SerializeObject(itemPayload));
            //Specify the content type. 
            content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            HttpResponseMessage result = client.PostAsync(resourceUrl, content).Result;
            if (result.IsSuccessStatusCode)
            {
                //email send successfully.
                Console.WriteLine("Email sent successfully. ");
            }
            else
            {
                //email send failed. check the result for detail information from REST api.
                Console.WriteLine("Email sent failed. Error: {0}", result.Content.ReadAsStringAsync().Result);
            }
        }

        //get microsoft graph
        private static string certificateAuthentication_graph()
        {
            string tenantId = "tesla329.onmicrosoft.com";
            string clientId = "dd784c4b-2956-44ae-9974-d84caf0e3d96";
            string resourceId = "https://graph.microsoft.com/";
            string resourceUrl = "https://graph.microsoft.com/v1.0/users/warpbubble@tesla329.onmicrosoft.com/sendmail"; //this is your on-behalf user's UPN
            string authority = String.Format("https://login.windows.net/{0}", tenantId);
            string certficatePath = @"C:\Dexter\Practice\outlook_mail\warpbubble.pfx"; //this is your certficate location.
            string certificatePassword = "p@ssw0rd"; // this is your certificate password
            X509Certificate2 certificate = new X509Certificate2(certficatePath, certificatePassword, X509KeyStorageFlags.MachineKeySet);
            AuthenticationContext authenticationContext = new AuthenticationContext(authority, false);

            ClientAssertionCertificate cac = new ClientAssertionCertificate(clientId, certificate);

            //get the access token to Outlook using the ClientAssertionCertificate
            var authenticationResult = authenticationContext.AcquireTokenAsync(resourceId, cac).Result;
            string token = authenticationResult.AccessToken;

            //var itemPayload = new
            //{
            //    Message = new
            //    {
            //        Subject = "Test email",
            //        Body = new { ContentType = "Text", Content = "this is test email." },
            //        ToRecipients = new[] { new { EmailAddress = new { Address = "user1@tesla329.onmicrosoft.com" } } }
            //    }
            //};

            //HttpClient client = new HttpClient();
            //client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
            //client.DefaultRequestHeaders.Add("Accept", "application/json");
            //HttpContent content = new StringContent(JsonConvert.SerializeObject(itemPayload));
            ////Specify the content type. 
            //content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            //HttpResponseMessage result = client.PostAsync(resourceUrl, content).Result;
            //if (result.IsSuccessStatusCode)
            //{
            //    //email send successfully.
            //    Console.WriteLine("Email sent successfully. ");
            //}
            //else
            //{
            //    //email send failed. check the result for detail information from REST api.
            //    Console.WriteLine("Email sent failed. Error: {0}", result.Content.ReadAsStringAsync().Result);
            //}
            return token;
        }

        private static void GetEvent(string Token)
        {
            var client = new RestClient("https://graph.microsoft.com/");
            var request = new RestRequest("/v1.0/users/warpbubble@tesla329.onmicrosoft.com/Events", Method.GET);
            request.AddHeader("Authorization", "Bearer " + Token);
            request.AddHeader("Content-Type", "application/json");
            request.AddHeader("Accept", "application/json");

            var response = client.Execute(request);
            var content = response.Content;
        }
        private static void GetMailMessage(string Token)
        {
            var client = new RestClient("https://graph.microsoft.com/");
            var request = new RestRequest("/v1.0/users/user1@tesla329.onmicrosoft.com/messages", Method.GET);
            request.AddHeader("Authorization", "Bearer " + Token);
            request.AddHeader("Content-Type", "application/json");
            request.AddHeader("Accept", "application/json");

            var response = client.Execute(request);
            var content = response.Content;
        }

        private static void creategraphclient(string token)
        {
            var graphserviceClient = new GraphServiceClient(
      new DelegateAuthenticationProvider(
          (requestMessage) =>
          {
              requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);

              return Task.FromResult(0);
          }));
            var messages = graphserviceClient.Users["user1@tesla329.onmicrosoft.com"].Messages.Request();
        }
        //https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/docs/overview.md
        private static void CreateAzuregraphClient(string token)
        {
            ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(new Uri("https://graph.windows.net/tesla329.onmicrosoft.com"),
                async () => { return await GetAccessToken(); });
            var user = activeDirectoryClient.Users["user1@tesla329.onmicrosoft.com"];

        }
        private static async Task<string> GetAccessToken()
        {
            string clientID = "a484dd2f-c252-427a-a021-cfc83ad9ad6a";  // The Client ID that we retrieved from the Azure Applications portal
            string key = "bdOB0aE9C9Rma5boXHxLYzCboBZ2W0x0O2a5uCILXLs=";  // The Client Key that we generated in the Azure Applications portal

            AuthenticationContext context = new AuthenticationContext("https://login.windows.net/" + "tesla329.onmicrosoft.com" + "/oauth2/token");

            ClientCredential credential = new ClientCredential(clientID, key);

            AuthenticationResult token = context.AcquireTokenAsync("00000002-0000-0000-c000-000000000000/graph.windows.net@" + "tesla329.onmicrosoft.com", credential).Result;

            return token.AccessToken;
        }
        // load from store
        //private static X509Certificate2 GetCertificate()
        //{
        //    X509Certificate2 certificate = null;
        //    var certStore = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        //    certStore.Open(OpenFlags.ReadOnly);
        //    var certCollection = certStore.Certificates.Find(X509FindType.FindByThumbprint, Thumbprint, false);
        //    // Get the first cert with the thumbprint
        //    if (certCollection.Count > 0)
        //    {
        //        certificate = certCollection[0];
        //    }
        //    certStore.Close();
        //    return certificate;
        //}
    }
}
//https://blog.appliedis.com/2016/07/28/office-365-and-the-graph-api-under-the-hood/



