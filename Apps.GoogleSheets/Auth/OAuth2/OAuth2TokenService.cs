using System.Text.Json;
using Apps.GoogleSheets.Constants;
using Blackbird.Applications.Sdk.Common.Authentication.OAuth2;

namespace Apps.GoogleSheets.Auth.OAuth2
{
    public class OAuth2TokenService : IOAuth2TokenService
    {
        private const string ExpiresAtKeyName = "expires_at";

        public bool IsRefreshToken(Dictionary<string, string> values)
            => DateTime.UtcNow > DateTime.Parse(values[ExpiresAtKeyName]);

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
            var bodyParameters = new Dictionary<string, string>
            {
                { "token", values["access_token"] },
            };

            return SendTokenRequest(Urls.Revoke, bodyParameters, CancellationToken.None);
        }

        private async Task<Dictionary<string, string>> RequestToken(Dictionary<string, string> bodyParameters,
            CancellationToken cancellationToken)
        {
            var tokenResponse = await SendTokenRequest(Urls.Token, bodyParameters, cancellationToken);
            var tokenData = await GetTokenData(tokenResponse);

            var resultDictionary = tokenData
                .ToDictionary(r => r.Key, r => r.Value?.ToString());

            var expiresIn = int.Parse(resultDictionary["expires_in"]);
            var expiresAt = DateTime.UtcNow.AddSeconds(expiresIn);

            resultDictionary.Add(ExpiresAtKeyName, expiresAt.ToString());
            return resultDictionary;
        }

        private async Task<Dictionary<string, object>> GetTokenData(HttpResponseMessage response)
        {
            var responseContent = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<Dictionary<string, object>>(responseContent)
                   ?? throw new InvalidOperationException($"Invalid response content: {responseContent}");
        }

        private Task<HttpResponseMessage> SendTokenRequest(
            string url,
            Dictionary<string, string> bodyParameters,
            CancellationToken cancellationToken)
        {
            using var httpClient = new HttpClient();
            using var httpContent = new FormUrlEncodedContent(bodyParameters);

            return httpClient.PostAsync(url, httpContent, cancellationToken);
        }
    }
}