using Blackbird.Applications.Sdk.Common.Authentication;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets
{
    public class GoogleDriveClient : DriveService
    {
        private static Initializer GetInitializer(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders)
        {
            var accessToken = authenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value;
            GoogleCredential credentials = GoogleCredential.FromAccessToken(accessToken);

            return new Initializer
            {
                HttpClientInitializer = credentials,
                ApplicationName = "Blackbird"
            };

        }

        public GoogleDriveClient(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders) : base(GetInitializer(authenticationCredentialsProviders)) { }
    }
}
