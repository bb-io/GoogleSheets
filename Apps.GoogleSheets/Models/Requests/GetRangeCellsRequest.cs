using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class GetRangeCellsRequest
    {
        [Display("Spread sheet ID")]
        public string SpreadSheetId { get; set; }

        [Display("Sheet name")]
        public string SheetName { get; set; }

        [Display("A Row ID")]
        public string RowIdA { get; set; }

        [Display("A Column")]
        public string ColumnA { get; set; }

        [Display("B Row ID")]
        public string RowIdB { get; set; }

        [Display("B Column")]
        public string ColumnB { get; set; }
    }
}
