using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.Models.Requests;

public class GetSpreadsheetsRequest
{
    [DataSource(typeof(FolderDataSourceHandler))]
    public string? FolderId { get; set; }
}
