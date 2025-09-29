using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.Models.Requests
{
    public class SheetRequest
    {
        [Display("Sheet title")]
        [DataSource(typeof(SheetDataSourceHandler))]
        public string SheetName { get; set; }
    }
}
