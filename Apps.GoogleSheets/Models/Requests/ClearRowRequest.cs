using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class ClearRowRequest
    {
        [Display("Row")] public string RowId { get; set; }
    }
}