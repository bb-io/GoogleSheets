using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            var filesListr = client.Files.List();
            var query = "mimeType='application/vnd.google-apps.spreadsheet' and trashed = false";
            if (context.SearchString != null)
            {
                query += $" and name contains '${context.SearchString}'";
            }
            filesListr.Q = query;
            filesListr.PageSize = 20;
            var filesList = filesListr.Execute();


            return filesList.Files.ToDictionary(x => x.Id, x => x.Name);
        }
    }
}