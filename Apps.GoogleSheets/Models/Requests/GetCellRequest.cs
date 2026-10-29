using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Exceptions;

namespace Apps.GoogleSheets.Models.Requests;

public class GetCellRequest
{
    [Display("Column", Description = "Column address (e.g. \"A\", \"B\", \"C\")")]
    public string Column { get; set; } = string.Empty;

    [Display("Row", Description = "Row number (e.g. \"1\", \"2\", \"3\")")]
    public string Row { get; set; } = string.Empty;

    public override string ToString()
    {
        if (!int.TryParse(Row, out int rowNumber))
            throw new PluginMisconfigurationException("The row value should be a number, e.g. 1");
        
        return $"{Column.ToUpperInvariant()}{rowNumber}";
    }
}