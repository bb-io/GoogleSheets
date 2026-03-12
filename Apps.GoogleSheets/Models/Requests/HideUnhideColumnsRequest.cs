using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class HideUnhideColumnsRequest
    {
        [Display("All columns", Description = "Hide or unhide all columns in the selected sheet")]
        public bool? AllColumns { get; set; }

        [Display("Start column", Description = "Column address, e.g. A, B, AA")]
        public string? StartColumn { get; set; }

        [Display("End column", Description = "Optional end column address, e.g. C, AB")]
        public string? EndColumn { get; set; }
    }
}
