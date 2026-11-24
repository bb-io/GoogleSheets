using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class InsertRowRequest
    {
        [Display("Row", Description = "Cell values in order. Use <EMPTY> to leave a cell blank")]
        public List<string> Row { get; set; }

        [Display("Start column address", Description = "Column address (e.g. \"A\", \"B\", \"C\")")]
        public string? ColumnAddress { get; set; }
    }
}
