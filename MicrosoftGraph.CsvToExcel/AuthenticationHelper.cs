using System;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace MicrosoftGraph.CsvToExcel
{
    public class AuthenticationHelper
    {
        // The Client ID is used by the application to uniquely identify itself to the v2.0 authentication endpoint.
        private const string ClientId = "REPLACE THIS VALUE WITH YOUR APPLICATION ID";
        public static string[] Scopes = { "Files.ReadWrite" };

        public static PublicClientApplication IdentityClientApp = new PublicClientApplication(ClientId);

        public static string TokenForUser;
        public static DateTimeOffset Expiration;

        private static GraphServiceClient _graphClient;

        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static GraphServiceClient GetAuthenticatedClient()
        {
            if (_graphClient != null)
                return _graphClient;

            try
            {
                _graphClient = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                    new DelegateAuthenticationProvider
                    (
                        async requestMessage =>
                        {
                            var token = await GetTokenForUserAsync();
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                        }
                    ));

                return _graphClient;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not create a graph client: " + ex.Message);
            }

            return _graphClient;
        }

        private static async Task<string> GetTokenForUserAsync()
        {
            AuthenticationResult authResult;

            try
            {
                authResult = await IdentityClientApp.AcquireTokenSilentAsync(Scopes, IdentityClientApp.Users.First());
                TokenForUser = authResult.AccessToken;
            }
            catch (Exception)
            {
                if (TokenForUser != null && Expiration > DateTimeOffset.UtcNow.AddMinutes(5))
                    return TokenForUser;

                authResult = await IdentityClientApp.AcquireTokenAsync(Scopes);
                TokenForUser = authResult.AccessToken;
                Expiration = authResult.ExpiresOn;
            }

            return TokenForUser;
        }

        public static void SignOut()
        {
            foreach (var user in IdentityClientApp.Users)
            {
                IdentityClientApp.Remove(user);
            }

            _graphClient = null;
            TokenForUser = null;
        }
    }
}