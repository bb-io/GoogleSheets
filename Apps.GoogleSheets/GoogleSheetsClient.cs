using Apps.GoogleSheets.Extensions.RateLimit;
using Blackbird.Applications.Sdk.Common.Authentication;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;

namespace Apps.GoogleSheets
{
    public class GoogleSheetsClient : SheetsService
    {
        private static Initializer GetInitializer(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders)
        {
            //var serviceAccountConfString = authenticationCredentialsProviders.First(p => p.KeyName == "serviceAccountConfString").Value;
            //string[] scopes = { SheetsService.Scope.Spreadsheets };
            //ServiceAccountCredential? credential = GoogleCredential.FromJson(serviceAccountConfString)
            //                                      .CreateScoped(scopes)
            //                                      .UnderlyingCredential as ServiceAccountCredential;
            var accessToken = authenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value;
            var credentials = GoogleCredential.FromAccessToken(accessToken);

            return new Initializer
            {
                HttpClientInitializer = credentials,
                ApplicationName = "Blackbird",
                HttpClientFactory = new RateLimitedHttpClientFactory()
            };
        }

        public GoogleSheetsClient(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders) : base(GetInitializer(authenticationCredentialsProviders)) 
        {
            this.HttpClient.Timeout = TimeSpan.FromSeconds(300);
        }
    }
}
