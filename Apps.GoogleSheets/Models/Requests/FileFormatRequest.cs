using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dictionaries;

namespace Apps.GoogleSheets.Models.Requests
{
    public class FileFormatRequest
    {

        [Display("File format")]
        [StaticDataSource(typeof(FileFormatDataSourceHandler))]
        public string Format { get; set; }
    }
}
