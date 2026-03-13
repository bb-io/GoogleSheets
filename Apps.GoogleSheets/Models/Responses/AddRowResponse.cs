using Apps.GoogleSheets.Dtos;
using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Responses;

public class AddRowResponse
{
    [Display("Row index")]
    public string RowId { get; set; }
    public RowDto Cells { get; set; }
}
