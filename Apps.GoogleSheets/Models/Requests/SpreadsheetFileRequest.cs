using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.GoogleSheets.Models.Requests
{
    public class SpreadsheetFileRequest
    {
        [Display("Spreadsheet ID")]
        [FileDataSource(typeof(SpreadsheetFilePickerDataSourceHandler))]
        public string SpreadSheetId { get; set; }
    }
}
