using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.GoogleSheets.Models.Requests;

public class GetSpreadsheetsRequest
{
    [FileDataSource(typeof(FolderPickerDataSourceHandler))]
    [Display("Folder ID")]
    public string? FolderId { get; set; }

    [Display("Get deleted spreadsheets")]
    public bool? FetchDeleted { get; set; }
}
