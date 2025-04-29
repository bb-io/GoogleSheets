using System.Net;

namespace Apps.GoogleSheets.Extensions.RateLimit
{
    public class RateLimitHandler : DelegatingHandler
    {
        private const int MaxRetryAttempts = 5;

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            int retryAttempts = 0;
            HttpResponseMessage response = await base.SendAsync(request, cancellationToken);

            while (response.StatusCode == (HttpStatusCode)429 && retryAttempts < MaxRetryAttempts)
            {
                if (response.Headers.TryGetValues("Retry-After", out var values) &&
                int.TryParse(System.Linq.Enumerable.FirstOrDefault(values), out int retryAfterSeconds))
                {
                    await Task.Delay(TimeSpan.FromSeconds(retryAfterSeconds), cancellationToken);
                }
                else
                {
                    await Task.Delay(TimeSpan.FromSeconds(Math.Pow(2, retryAttempts)), cancellationToken);
                }

                retryAttempts++;
                response = await base.SendAsync(request, cancellationToken);
            }

            return response;
        }
    }
}
