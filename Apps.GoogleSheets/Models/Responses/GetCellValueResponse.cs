using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Responses
{
    public class GetCellValueResponse
    {
        [Display("Cell value")]
        public string CellValue { get; set; }
    }
}
