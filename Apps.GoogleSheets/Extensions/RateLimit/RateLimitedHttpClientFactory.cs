using Google.Apis.Http;

namespace Apps.GoogleSheets.Extensions.RateLimit
{
    public class RateLimitedHttpClientFactory : Google.Apis.Http.IHttpClientFactory
    {
        public ConfigurableHttpClient CreateHttpClient(CreateHttpClientArgs args)
        {
            HttpMessageHandler innerHandler = new HttpClientHandler();

            var rateLimitHandler = new RateLimitHandler
            {
                InnerHandler = innerHandler
            };

            var configurableMessageHandler = new ConfigurableMessageHandler(rateLimitHandler);

            foreach (var initializer in args.Initializers)
            {
                if (initializer is IHttpExecuteInterceptor interceptor)
                {
                    configurableMessageHandler.AddExecuteInterceptor(interceptor);
                }
            }

            var configurableHttpClient = new ConfigurableHttpClient(configurableMessageHandler)
            {
                Timeout = TimeSpan.FromSeconds(300)
            };

            foreach (var initializer in args.Initializers)
            {
                initializer.Initialize(configurableHttpClient);
            }

            return configurableHttpClient;
        }
    }
}
