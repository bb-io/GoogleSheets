using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Connections;

namespace Apps.GoogleSheets.Connections
{
    public class ConnectionDefinition : IConnectionDefinition
    {
        public IEnumerable<ConnectionPropertyGroup> ConnectionPropertyGroups => new List<ConnectionPropertyGroup>
        {
            new()
            {
                Name = "Oauth",
                AuthenticationType = ConnectionAuthenticationType.OAuth2,
                ConnectionUsage = ConnectionUsage.Actions,
                ConnectionProperties = new List<ConnectionProperty>()
            }
            //new ConnectionPropertyGroup
            //{
            //    Name = "Service account",
            //    AuthenticationType = ConnectionAuthenticationType.Undefined,
            //    ConnectionUsage = ConnectionUsage.Actions,
            //    ConnectionProperties = new List<ConnectionProperty>()
            //    {
            //        new ConnectionProperty("serviceAccountConfString")
            //    }
            //}
        };

        public IEnumerable<AuthenticationCredentialsProvider> CreateAuthorizationCredentialsProviders(
            Dictionary<string, string> values)
        {
            var accessToken = values.First(v => v.Key == "access_token");
            yield return new AuthenticationCredentialsProvider(
                AuthenticationCredentialsRequestLocation.None,
                "Authorization",
                accessToken.Value
            );
            //var serviceAccountConfString = values.First(v => v.Key == "serviceAccountConfString");
            //yield return new AuthenticationCredentialsProvider(
            //    AuthenticationCredentialsRequestLocation.None,
            //    serviceAccountConfString.Key,
            //    serviceAccountConfString.Value
            //);
        }
    }
}