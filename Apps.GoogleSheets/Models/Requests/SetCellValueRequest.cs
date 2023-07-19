using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class SetCellValueRequest
    {
        [Display("Spead sheet ID")] public string SpreadSheetId { get; set; }
        [Display("Sheet name")] public string SheetName { get; set; }
        [Display("Row ID")] public string RowId { get; set; }
        [Display("Column")] public string Column { get; set; }
        [Display("Value")] public string Value { get; set; }
    }
}