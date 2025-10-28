using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests;

public class ClearRangeRequest
{
    [Display("Range", Description = "Range to clear, e.g., A1:B10 or C5:C10.")]
    public string Range { get; set; }
}
