using Apps.GoogleSheets.Dtos;

namespace Apps.GoogleSheets.Models.Responses
{
    public class GetRangeCellsResponse
    {
        public IEnumerable<SheetRowDto> Rows { get; set; }
    }
}
