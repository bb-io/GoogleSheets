using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.Models.Requests;

public class GetSpreadsheetsRequest
{
    [DataSource(typeof(FolderDataSourceHandler))]
    [Display("Folder ID")]
    public string? FolderId { get; set; }
}
