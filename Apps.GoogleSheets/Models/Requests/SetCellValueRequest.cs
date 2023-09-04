using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.Models.Requests
{
    public class SetCellValueRequest
    {
        [DataSource(typeof(SpreadsheetsDataHandler))]
        [Display("Speadsheet")] public string SpreadSheetId { get; set; }
        [Display("Sheet name")] public string SheetName { get; set; }
        [Display("Cell")] public string CellId { get; set; }
        [Display("Value")] public string Value { get; set; }
    }
}