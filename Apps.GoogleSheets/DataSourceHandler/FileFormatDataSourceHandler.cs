using Blackbird.Applications.Sdk.Common.Dictionaries;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.DataSourceHandler;

public class FileFormatDataSourceHandler : IStaticDataSourceItemHandler
{
    public IEnumerable<DataSourceItem> GetData()
    {
        return new List<DataSourceItem>
    {
        new("PDF", ".pdf"),
        new("CSV", ".csv"),
        new("XLSX", ".xlsx"),
    };

    }
}