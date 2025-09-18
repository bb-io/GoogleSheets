using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;

namespace Apps.GoogleSheets.DataSourceHandler;

public class FolderDataSourceHandler : BaseInvocable, IDataSourceItemHandler
{
    private IEnumerable<AuthenticationCredentialsProvider> Creds =>
        InvocationContext.AuthenticationCredentialsProviders;

    public FolderDataSourceHandler(InvocationContext invocationContext) : base(invocationContext)
    {
    }

    public IEnumerable<DataSourceItem> GetData(DataSourceContext context)
    {
        var client = new GoogleDriveClient(Creds);
        var query = "mimeType = 'application/vnd.google-apps.folder'";
        if (context.SearchString != null)
            query += $" and name contains '{context.SearchString}'";

        var filesListr = client.Files.List();

        filesListr.IncludeItemsFromAllDrives = true;
        filesListr.SupportsAllDrives = true;
        filesListr.Q = query;
        filesListr.PageSize = 20;

        var filesList = filesListr.Execute();

        return filesList.Files.Select(x => new DataSourceItem(x.Id, x.Name));
    }
}
