using Apps.GoogleSheets.Models.Requests;
using Apps.GoogleSheets.Polling;
using Apps.GoogleSheets.Polling.Models;
using Blackbird.Applications.Sdk.Common.Polling;
using Tests.GoogleSheets.Base;

namespace Tests.GoogleSheets;

[TestClass]
public class CellPollingTests : TestBase
{
    private CellPollingList Polling => new(InvocationContext);
    
    [TestMethod]
    public async Task OnValueAddedToCell_IsSuccess()
    {
        // Arrange
        var memory = new CellValueMemory { LastValue = "jjj" };
        var request = new PollingEventRequest<CellValueMemory> { Memory = memory };
        var spreadsheet = new SpreadsheetFileRequest { SpreadSheetId = "1YenaAZWzJzcmjsHNdyolxVQdudrCHIRlF8MR0sv2-CE" };
        var sheet = new SheetRequest { SheetName = "Sheet1" };
        var cellAddress = new GetCellRequest { Column = "A", Row = "2" };

        // Act
        var result = await Polling.OnValueAddedToCell(request, spreadsheet, sheet, cellAddress);

        // Assert
        PrintJsonResult(result);
    }
}