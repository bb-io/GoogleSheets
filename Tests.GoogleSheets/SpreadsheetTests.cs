using Apps.GoogleSheets.Actions;
using Apps.GoogleSheets.Models.Requests;
using Tests.GoogleSheets.Base;

namespace Tests.GoogleSheets
{
    [TestClass]
    public class SpreadsheetTests : TestBase
    {
        [TestMethod]
        public async Task GetRange_ReturnsSuccess()
        {
            var action = new SpreadsheetActions(InvocationContext, FileManager);

            var spreadsheetFileRequest = new SpreadsheetFileRequest { SpreadSheetId= "" };
            var spreadSheet = new SheetRequest { SheetName= "YOUR_FILE_NAME" };
            var rangeRequest = new RangeRequest {StartCell= "A2", EndCell= "C3" };

            var result = await action.GetRange(spreadsheetFileRequest, spreadSheet, rangeRequest);

            Assert.IsNotNull(result);
        }

    }
}
