using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.Models.Requests
{
    public class UpdateCellRequest
    {
        [Display("Value", Description = "Cell value")]
        public string Value { get; set; }
    }
}