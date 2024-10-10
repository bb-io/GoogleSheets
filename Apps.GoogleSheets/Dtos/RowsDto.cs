

using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Dtos
{
    public class RowsDto
    {
        public List<List<string>> Rows { get; set; }

        [Display("Rows count")]
        public double RowsCount { get; set; }

        [Display("Row numbers")]
        public IEnumerable<int> RowIds {get; set;}
    }
}
