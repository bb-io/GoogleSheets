using Apps.GoogleSheets.DataSourceHandler;
using Apps.GoogleSheets.Models.Requests;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Tests.GoogleSheets.Base;

namespace Tests.GoogleSheets;

[TestClass]
public class DatahandlerTests :TestBase
{
    [TestMethod]
    public async Task SpreadsheetDatahandler_IsSuccess()
    {
        var handler = new SpreadsheetFileDataSourceHandler(InvocationContext);

        var result = handler.GetData(new DataSourceContext { });

        Console.WriteLine($"Total: {result.Count()}");
        foreach (var item in result)
        {
            Console.WriteLine($"{item.Value}: {item.Key}");
        }

        Assert.IsTrue(result.Count() > 0);
    }

    [TestMethod]
    public async Task SheetDatahandler_IsSuccess()
    {
        var handler = new SheetDataSourceHandler(InvocationContext, new SpreadsheetFileRequest {SpreadSheetId= "" });

        var result = await handler.GetDataAsync(new DataSourceContext { }, CancellationToken.None);

        Console.WriteLine($"Total: {result.Count()}");
        foreach (var item in result)
        {
            Console.WriteLine($"{item.Value}: {item.Key}");
        }

        Assert.IsTrue(result.Count() > 0);
    }

    [TestMethod]
    public void FolderDataHandler_IsSuccess()
    {
        // Arrange
        var handler = new FolderDataSourceHandler(InvocationContext);

        // Act
        var result = handler.GetData(new DataSourceContext { });

        // Assert
        Console.WriteLine($"Total: {result.Count()}");
        foreach (var item in result)
        {
            Console.WriteLine($"{item.Value}: {item.DisplayName}");
        }
    }
}
