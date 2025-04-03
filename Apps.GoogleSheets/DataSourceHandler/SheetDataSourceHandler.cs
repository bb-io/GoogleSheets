using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Invocation;
using Apps.GoogleSheets.Models.Requests;
using Google.Apis.Sheets.v4.Data;

namespace Apps.GoogleSheets.DataSourceHandler
{
    public class SheetDataSourceHandler : BaseInvocable, IAsyncDataSourceHandler
    {
        private SpreadsheetFileRequest SpreadsheetFileRequest { get; set; }

        private IEnumerable<AuthenticationCredentialsProvider> Creds =>
            InvocationContext.AuthenticationCredentialsProviders;

        public SheetDataSourceHandler(InvocationContext invocationContext, [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest) : base(invocationContext)
        {
            SpreadsheetFileRequest = spreadsheetFileRequest;
        }

        public async Task<Dictionary<string, string>> GetDataAsync(DataSourceContext context, CancellationToken cancellationToken)
        {
            var client = new GoogleSheetsClient(Creds);
            var spreadsheet = await client.Spreadsheets.Get(SpreadsheetFileRequest.SpreadSheetId).ExecuteAsync();

            var sheets = new List<Sheet>();
            if (context.SearchString != null)
            {
                sheets = spreadsheet.Sheets.Where(x => x.Properties.Title.Contains(context.SearchString, StringComparison.InvariantCultureIgnoreCase)).ToList();
            }
            else
            {
                sheets = spreadsheet.Sheets.ToList();
            }
            return sheets.ToDictionary(x => x.Properties.Title, x => x.Properties.Title);
        }
    }
}
