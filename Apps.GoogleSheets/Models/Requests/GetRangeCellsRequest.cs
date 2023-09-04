using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.Models.Requests
{
    public class GetRangeCellsRequest
    {
        [DataSource(typeof(SpreadsheetsDataHandler))]
        [Display("Spreadsheet")]
        public string SpreadSheetId { get; set; }

        [Display("Sheet name")]
        public string SheetName { get; set; }

        [Display("Cell start")]
        public string CellStart { get; set; }

        [Display("Cell end")]
        public string CellEnd { get; set; }
    }
}
