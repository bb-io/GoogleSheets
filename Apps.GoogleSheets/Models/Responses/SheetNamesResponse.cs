using Apps.GoogleSheets.Models.Dto;
using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Responses
{
    public class SheetNamesResponse
    {
        [Display("Sheets")]
        public List<SheetNameDto> Sheets { get; set; } = new();
    }
}
