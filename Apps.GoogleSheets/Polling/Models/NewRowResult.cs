using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Polling.Models
{
    public class NewRowResult
    {
        [Display("New rows")]
        public List<NewRow> NewRows { get; set; }
    }

    public class NewRow
    {
        public int RowIndex { get; set; }
        public List<string> RowValues { get; set; }
    }
}
