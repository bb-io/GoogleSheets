using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.GoogleSheets.Models.Requests;

public class CopySpreadsheetRequest
{
    [Display("New spreadsheet name")]
    public string NewSpreadsheetName { get; set; }

    [Display("Folder ID", Description = "Specify the folder where the new spreadsheet should be copied.")]
    [FileDataSource(typeof(FolderPickerDataSourceHandler))]
    public string? FolderId { get; set; }
}
