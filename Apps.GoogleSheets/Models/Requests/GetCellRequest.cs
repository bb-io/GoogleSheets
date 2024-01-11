using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.Models.Requests
{
    public class GetCellRequest
    {
        [Display("Cell address", Description = "Cell address (e.g. \"A1\", \"B2\", \"C3\")")]
        public string Cell { get; set; }
    }
}