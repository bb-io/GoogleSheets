using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.GoogleSheets.Models.Requests
{
    public class MoveFileRequest
    {
        [Display("File ID")]
        [FileDataSource(typeof(SpreadsheetFilePickerDataSourceHandler))]
        public string FileId { get; set; } = default!;

        [Display("New parent folder ID")]
        [FileDataSource(typeof(FolderPickerDataSourceHandler))]
        public string NewParentFolderId { get; set; } = default!;

        [Display("Move target if shortcut", Description = "If the file is a Google Drive shortcut, move the target file instead of the shortcut itself.")]
        public bool? MoveTargetIfShortcut { get; set; }
    }
}
