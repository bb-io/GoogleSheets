using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.GoogleSheets.Models.Requests;

public class CreateSpreadsheetRequest
{
    [Display("Title")]
    public string Title { get; set; }

    [Display("Sheet name")]
    public string? InitialSheetName { get; set; }

    [Display("Folder ID", Description = "Specify the folder where you want to save your spreadsheet.")]
    [FileDataSource(typeof(FolderPickerDataSourceHandler))]
    public string? FolderId { get; set; }
}
