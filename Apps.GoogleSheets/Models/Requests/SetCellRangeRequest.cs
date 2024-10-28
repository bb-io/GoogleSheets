using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class SetCellRangeRequest
    {

        [Display("Cell start")]
        public string CellStart { get; set; }

        [Display("Cell end")]
        public string CellEnd { get; set; }
        [Display("Values")] public IEnumerable<string> Columns { get; set; }
    }
}