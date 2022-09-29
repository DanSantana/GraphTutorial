using Azure.Core;
using Azure.Identity;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Graph;

namespace GraphTutorial
{
    public class GraphHelper
    {
        // Settings object
        private static Settings? _settings;
        // User auth token credential
        private static DeviceCodeCredential? _deviceCodeCredential;
        // Client configured with user authentication
        private static GraphServiceClient? _userClient;

        public static void InitializeGraphForUserAuth(Settings settings,
            Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
        {
            _settings = settings;

            _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
                settings.AuthTenant, settings.ClientId);

            _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
        }

        public static async Task<string> GetUserTokenAsync()
        {
            // Ensure credential isn't null
            _ = _deviceCodeCredential ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            // Ensure scopes isn't null
            _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

            // Request token with given scopes
            var context = new TokenRequestContext(_settings.GraphUserScopes);
            var response = await _deviceCodeCredential.GetTokenAsync(context);
            return response.Token;
        }
        public static Task<Microsoft.Graph.User> GetUserAsync()
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            return _userClient.Me
                .Request()
                .Select(u => new
                {
            // Only request specific properties
            u.DisplayName,
                    u.Mail,
                    u.UserPrincipalName
                })
                .GetAsync();
        }

        public static Task<IMailFolderMessagesCollectionPage> GetInboxAsync()
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            return _userClient.Me
                // Only messages from Inbox folder
                .MailFolders["Inbox"]
                .Messages
                .Request()
                .Select(m => new
                {
            // Only request specific properties
            m.From,
                    m.IsRead,
                    m.ReceivedDateTime,
                    m.Subject
                })
                // Get at most 25 results
                .Top(25)
                // Sort by received time, newest first
                .OrderBy("ReceivedDateTime DESC")
                .GetAsync();
        }




    }
}
