using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Dto
{
    public class SheetNameDto
    {
        [Display("Sheet name")]
        public string Name { get; set; } = default!;

        [Display("Sheet ID")]
        public int? SheetId { get; set; }
    }
}
