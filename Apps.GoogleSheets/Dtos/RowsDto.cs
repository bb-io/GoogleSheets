

using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Dtos
{
    public class RowsDto
    {
        public List<List<string>> Rows { get; set; }

        [Display("Rows Count")]
        public double RowsCount { get; set; }
    }
}
