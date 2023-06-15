using Blackbird.Applications.Sdk.Common.Authentication.OAuth2;
using System.Text.Json;

namespace Apps.GoogleSheets.Authorization.OAuth2
{
    public class OAuth2TokenService : IOAuth2TokenService
    {
        private const string ExpiresAtKeyName = "expires_at";

        public bool IsRefreshToken(Dictionary<string, string> values)
        {
            return false;
        }

        public async Task<Dictionary<string, string>> RefreshToken(Dictionary<string, string> values, CancellationToken cancellationToken)
        {
            throw new NotImplementedException();
        }

        public async Task<Dictionary<string, string?>> RequestToken(
            string state, 
            string code, 
            Dictionary<string, string> values, 
            CancellationToken cancellationToken)
        {
            return new Dictionary<string, string?>() { { "access_token", values["access_token"] } };
        }

        public Task RevokeToken(Dictionary<string, string> values)
        {
            throw new NotImplementedException();
        }
    }
}
