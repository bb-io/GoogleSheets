using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.Models.Requests
{
    public class GetCellRequest
    {
        [Display("Column", Description = "Column address (e.g. \"A\", \"B\", \"C\")")]
        public string Column { get; set; }

        [Display("Row", Description = "Row number (e.g. \"1\", \"2\", \"3\")")]
        public int Row { get; set; }
    }
}