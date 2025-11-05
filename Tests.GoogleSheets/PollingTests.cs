using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Apps.GoogleSheets.Models.Requests;
using Apps.GoogleSheets.Polling;
using Apps.GoogleSheets.Polling.Models;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Blackbird.Applications.Sdk.Common.Polling;
using Tests.GoogleSheets.Base;

namespace Tests.GoogleSheets
{
    [TestClass]
    public class PollingTests : TestBase
    {
        [TestMethod]
        public async Task Polling_ReturnsSucces()
        {
            var spreadsheetId = "17ieaCd7SXacxaFr7LkhfdiFRVToyBz1kIzoi6IqM8oc";
            var sheetName = "Стальна шерсть";
            int currentRowCount = 6;

            var actions = new GoogleSheetsPollingList(InvocationContext);

            var spreadId = new SpreadsheetFileRequest
            {
                SpreadSheetId = spreadsheetId
            };
            var spreadName = new SheetRequest
            {
                SheetName = sheetName
            };

            var pollingRequest = new PollingEventRequest<NewRowAddedMemory>
            {
                Memory = new NewRowAddedMemory
                {
                    LastRowCount = currentRowCount,
                    LastPollingTime = DateTime.UtcNow,
                    Triggered = false
                }
            };

            var response = await actions.OnNewRowsAdded(pollingRequest, spreadId, spreadName);

            if (response.Result != null)
            {
                foreach (var result in response.Result)
                {
                    foreach (var newRow in result.NewRows)
                    {
                        Console.WriteLine($"Row Index: {newRow.RowIndex}");
                        foreach (var cellValue in newRow.RowValues)
                        {
                            Console.WriteLine($"   Cell: {cellValue}");
                        }
                    }
                }
            }

            Assert.IsNotNull(response, "Response is null.");
        }

        [TestMethod]
        public async Task OnNewRowsAdded_NonExistingSpreadsheet_ThrowsMisconfigException()
        {
            // Arrange
            var actions = new GoogleSheetsPollingList(InvocationContext);
            var spreadId = new SpreadsheetFileRequest { SpreadSheetId = "nonexisting" };
            var spreadName = new SheetRequest { SheetName = "Стальна шерсть" };

            var pollingRequest = new PollingEventRequest<NewRowAddedMemory>
            {
                Memory = new NewRowAddedMemory
                {
                    LastRowCount = 6,
                    LastPollingTime = DateTime.UtcNow,
                    Triggered = false
                }
            };
            // Act
            var ex = await Assert.ThrowsExceptionAsync<PluginMisconfigurationException>(async () => 
                await actions.OnNewRowsAdded(pollingRequest, spreadId, spreadName)
            );

            // Assert
            StringAssert.Contains(ex.Message, "was not found");
        }
    }
}
