using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.Models.Requests
{
    public class CreateSpreadsheetRequest
    {
        [Display("Title")]
        public string Title { get; set; }

        [Display("Sheet name")]
        public string? InitialSheetName { get; set; }
    }
}
