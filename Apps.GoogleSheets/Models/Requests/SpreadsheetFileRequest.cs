using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.Models.Requests
{
    public class SpreadsheetFileRequest
    {
        [Display("Spreadsheet ID")]
        [DataSource(typeof(SpreadsheetFileDataSourceHandler))]
        public string SpreadSheetId { get; set; }
    }
}
