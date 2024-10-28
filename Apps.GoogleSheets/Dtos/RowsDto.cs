

using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Dtos
{
    public class RowsDto
    {
        public List<_row> Rows { get; set; }

        [Display("Rows count")]
        public double RowsCount { get; set; }
    }

    public class _row
    {
        [Display("Row ID")]
        public int RowId { get; set; }
        public List<string> Values { get; set; }
    }
}
