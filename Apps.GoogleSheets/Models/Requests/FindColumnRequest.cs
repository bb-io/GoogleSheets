using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class FindColumnRequest
    {
        [Display("Row ID")]
        public string RowId { get; set; }

        [Display("Value")]
        public string Value { get; set; }
    }
}
