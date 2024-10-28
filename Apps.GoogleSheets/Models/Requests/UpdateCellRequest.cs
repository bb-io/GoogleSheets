using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class UpdateCellRequest
    {
        [Display("Value", Description = "Cell value")]
        public string Value { get; set; }
    }
}