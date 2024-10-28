using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class RangeRequest
    {
        [Display("Start Cell", Description = "Cell address (e.g. \"A1\", \"B2\", \"C3\")")]
        public string StartCell { get; set; }

        [Display("End cell", Description = "Cell address (e.g. \"A1\", \"B2\", \"C3\")")]
        public string EndCell { get; set; }
    }
}
