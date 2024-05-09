using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class FindRowRequest
    {
        [Display("Column", Description = "Column address (e.g. \"A\", \"B\", \"C\")")]
        public string Column { get; set; }

        public string Value { get; set; }
    }
}
