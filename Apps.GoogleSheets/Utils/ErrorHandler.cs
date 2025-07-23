using Blackbird.Applications.Sdk.Common.Exceptions;
using Google;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            catch (Exception ex)
            {
                throw new PluginApplicationException($"Error: {ex.Message} - {ex.InnerException}");
            }
        }
    }
}
