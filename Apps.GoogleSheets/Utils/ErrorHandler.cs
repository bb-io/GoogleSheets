using Blackbird.Applications.Sdk.Common.Exceptions;
using Google;

namespace Apps.GoogleSheets.Utils
{
    public static class ErrorHandler
    {
        public static async Task ExecuteWithErrorHandlingAsync(Func<Task> action)
        {
            try
            {
                await action();
            }
            catch (GoogleApiException ex)
            {
                throw new PluginMisconfigurationException($"Error: status {ex.HttpStatusCode}, {ex.Message}. Please check your inputs and try again");
            }
            catch (IndexOutOfRangeException ex)
            {
                throw new PluginApplicationException(
                    "Index was out of range. The requested range or returned data is outside the expected bounds.", ex);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                throw new PluginApplicationException(
                    "Argument was out of range. The requested range or parameters are invalid for the current sheet.", ex);
            }
            catch (Exception ex)
            {
                throw new PluginApplicationException($"Error: {ex.Message} - {ex.InnerException}");
            }
        }

        public static async Task<T> ExecuteWithErrorHandlingAsync<T>(Func<Task<T>> action)
        {
            try
            {
                return await action();
            }
            catch (GoogleApiException ex)
            {
                throw new PluginMisconfigurationException($"Error: status {ex.HttpStatusCode}, {ex.Message}. Please check your inputs and try again");
            }
            catch (IndexOutOfRangeException ex)
            {
                throw new PluginApplicationException(
                    "Index was out of range. The requested range or returned data is outside the expected bounds.", ex);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                throw new PluginApplicationException(
                    "Argument was out of range. The requested range or parameters are invalid for the current sheet.", ex);
            }
            catch (Exception ex)
            {
                throw new PluginApplicationException($"Error: {ex.Message} - {ex.InnerException}");
            }
        }
    }
}
