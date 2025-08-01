﻿using Apps.GoogleSheets.Actions;
using Apps.GoogleSheets.Models.Requests;
using Apps.GoogleSheets.Models.Responses;
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

            var spreadsheetFileRequest = new SpreadsheetFileRequest { SpreadSheetId= "17ieaCd7SXacxaFr7LkhfdiFRVToyBz1kIzoi6IqM8oc" };
            var spreadSheet = new SheetRequest { SheetName= "__" };
            var rangeRequest = new RangeRequest {StartCell= "A1", EndCell= "C3" };

            var result = await action.GetRange(spreadsheetFileRequest, spreadSheet, rangeRequest);

            foreach (var row in result.Rows)
            {
                Console.Write($"Row {row.RowId}: ");
                foreach (var value in row.Values)
                {
                    Console.Write($"{value}, ");
                }
                Console.WriteLine();
            }

            Assert.IsNotNull(result);
        }

        [TestMethod]
        public async Task DownloadCsv_with_range_ReturnsSuccess()
        {
            var action = new SpreadsheetActions(InvocationContext, FileManager);

            var spreadsheetFileRequest = new SpreadsheetFileRequest { SpreadSheetId = "" };
            var spreadSheet = new SheetRequest { SheetName = "YOUR_FILE_NAME" };
            var rangeRequest = new OptionalRangeRequest { StartCell = "B3", EndCell = "D17" };

            var result = await action.DownloadCSV(spreadsheetFileRequest, spreadSheet, rangeRequest, new CsvOptions { });

            Assert.IsNotNull(result);
        }

        [TestMethod]
        public async Task DownloadCsv_without_range_ReturnsSuccess()
        {
            var action = new SpreadsheetActions(InvocationContext, FileManager);

            var spreadsheetFileRequest = new SpreadsheetFileRequest { SpreadSheetId = "" };
            var spreadSheet = new SheetRequest { SheetName = "" };
            var rangeRequest = new OptionalRangeRequest {  };

            var result = await action.DownloadCSV(spreadsheetFileRequest, spreadSheet, rangeRequest, new CsvOptions { });

            Assert.IsNotNull(result);
        }

        [TestMethod]
        public async Task ImportCsv_with_range_ReturnsSuccess()
        {
            var action = new SpreadsheetActions(InvocationContext, FileManager);

            var spreadsheetFileRequest = new SpreadsheetFileRequest { SpreadSheetId = "17ieaCd7SXacxaFr7LkhfdiFRVToyBz1kIzoi6IqM8oc" };
            var spreadSheet = new SheetRequest { SheetName = "Стальна шерсть" };
            var rangeRequest = new OptionalRangeRequest
            {
              
            };
            var fileCsv = new FileResponse { File = new Blackbird.Applications.Sdk.Common.Files.FileReference { Name= "test.csv" } };

            var result = await action.ImportCSVAppend(spreadsheetFileRequest, spreadSheet, fileCsv, new CsvOptions { },null);

            Assert.IsNotNull(result);
        }


        [TestMethod]
        public async Task ExportGlossary_ReturnsSuccess()
        {
            var action = new SpreadsheetActions(InvocationContext, FileManager);

            var spreadsheetFileRequest = new SpreadsheetFileRequest { SpreadSheetId = "1zq3Ztl-dBUn_PWQS-J0FaiPGehw-nfqo1Arv_a7W42A" };
            var spreadSheet = new SheetRequest { SheetName = "Лист1" };
            var rangeRequest = new OptionalRangeRequest { };

            var result = await action.ExportGlossary(spreadsheetFileRequest, spreadSheet, string.Empty,string.Empty);

            Assert.IsNotNull(result);
        }

    }
}
