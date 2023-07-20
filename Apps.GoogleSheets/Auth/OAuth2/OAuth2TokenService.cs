using System.Text.Json;
using Apps.GoogleSheets.Constants;
using Blackbird.Applications.Sdk.Common.Authentication.OAuth2;

namespace Apps.GoogleSheets.Auth.OAuth2
{
    public class OAuth2TokenService : IOAuth2TokenService
    {
        private const string ExpiresAtKeyName = "expires_at";

        public bool IsRefreshToken(Dictionary<string, string> values)
            => values.TryGetValue(ExpiresAtKeyName, out var expireValue) &&
               DateTime.UtcNow > DateTime.Parse(expireValue);

        public Task<Dictionary<string, string>> RefreshToken(Dictionary<string, string> values,
            CancellationToken cancellationToken)
        {
            const string grant_type = "refresh_token";

            var bodyParameters = new Dictionary<string, string>
            {
                { "grant_type", grant_type },
                { "client_id", ApplicationConstants.ClientId },
                { "client_secret", ApplicationConstants.ClientSecret },
                { "refresh_token", values["refresh_token"] },
            };

            return RequestToken(bodyParameters, cancellationToken);
        }

        public Task<Dictionary<string, string>> RequestToken(
            string state,
            string code,
            Dictionary<string, string> values,
            CancellationToken cancellationToken)
        {
            const string grant_type = "authorization_code";

            var bodyParameters = new Dictionary<string, string>
            {
                { "grant_type", grant_type },
                { "client_id", ApplicationConstants.ClientId },
                { "client_secret", ApplicationConstants.ClientSecret },
                { "redirect_uri", ApplicationConstants.RedirectUri },
                { "code", code }
            };

            return RequestToken(bodyParameters, cancellationToken);
        }

        public Task RevokeToken(Dictionary<string, string> values)
        {
            throw new NotImplementedException();
        }

        private async Task<Dictionary<string, string>> RequestToken(Dictionary<string, string> bodyParameters,
            CancellationToken cancellationToken)
        {
            var utcNow = DateTime.UtcNow;
            using HttpClient httpClient = new HttpClient();
            using var httpContent = new FormUrlEncodedContent(bodyParameters);
            using var response = await httpClient.PostAsync(Urls.Token, httpContent, cancellationToken);
            var responseContent = await response.Content.ReadAsStringAsync();
            var resultDictionary = JsonSerializer.Deserialize<Dictionary<string, object>>(responseContent)
                                       ?.ToDictionary(r => r.Key, r => r.Value?.ToString())
                                   ?? throw new InvalidOperationException(
                                       $"Invalid response content: {responseContent}");
            var expiresIn = int.Parse(resultDictionary["expires_in"]);
            var expiresAt = utcNow.AddSeconds(expiresIn);
            resultDictionary.Add(ExpiresAtKeyName, expiresAt.ToString());
            return resultDictionary;
        }
    }
}