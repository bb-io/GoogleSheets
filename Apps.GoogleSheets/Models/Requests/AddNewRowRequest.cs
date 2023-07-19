using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class AddNewRowRequest
    {
        [Display("Spread sheet ID")] public string SpreadSheetId { get; set; }
        [Display("Sheet name")] public string SheetName { get; set; }
        [Display("Table start row ID")] public string TableStartRowId { get; set; }
        [Display("Table start column")] public string TableStartColumn { get; set; }
        [Display("Columns")] public IEnumerable<string> Columns { get; set; }
    }
}