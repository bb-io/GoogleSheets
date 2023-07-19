using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class ClearRowRequest
    {
        [Display("Spread sheet ID")] public string SpreadSheetId { get; set; }
        [Display("Sheet name")] public string SheetName { get; set; }
        [Display("Row ID")] public string RowId { get; set; }
    }
}