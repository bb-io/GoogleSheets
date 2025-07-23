using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common;
using Google.Apis.Sheets.v4.Data;

namespace Apps.GoogleSheets.DataSourceHandler
{
    public class SpreadsheetFileDataSourceHandler : BaseInvocable, IDataSourceHandler
    {
        private IEnumerable<AuthenticationCredentialsProvider> Creds =>
            InvocationContext.AuthenticationCredentialsProviders;

        public SpreadsheetFileDataSourceHandler(InvocationContext invocationContext) : base(invocationContext)
        {
        }

        public Dictionary<string, string> GetData(DataSourceContext context)
        {
            var client = new GoogleDriveClient(Creds);
            var request = client.Files.List();
            var query = "mimeType='application/vnd.google-apps.spreadsheet' and trashed = false";
            if (!string.IsNullOrEmpty(context.SearchString))
            {
                query += $" and name contains '{context.SearchString}'";
            }
            request.IncludeItemsFromAllDrives = true;
            request.SupportsAllDrives = true;
            request.Q = query;

            request.PageSize = 1000;

            var allFiles = new List<Google.Apis.Drive.v3.Data.File>();
            string pageToken = null;

            do
            {
                request.PageToken = pageToken;
                var response = request.Execute(); 

                if (response.Files != null && response.Files.Any())
                    allFiles.AddRange(response.Files);

                pageToken = response.NextPageToken;
            }
            while (!string.IsNullOrEmpty(pageToken));

            return allFiles
                .ToDictionary(x => x.Id, x => x.Name);
        }
    }
}