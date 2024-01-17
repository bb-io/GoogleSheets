using Blackbird.Applications.Sdk.Common;
using Google.Apis.Sheets.v4.Data;

namespace Apps.GoogleSheets.Dtos;

public class SheetDto
{
    public SheetDto() { }

    public SheetDto(SheetProperties sheetProperties)
    {
        SheetId = sheetProperties.SheetId.ToString()!;
        Title = sheetProperties.Title;
        Index = sheetProperties.Index!.Value;
        SheetType = sheetProperties.SheetType;
    }
    
    [Display("Sheet ID")]
    public string SheetId { get; set; }

    public string Title { get; set; }

    public int Index { get; set; }

    [Display("Sheet type")]
    public string SheetType { get; set; }
}
