using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Apps.GoogleSheets.Models.Requests;
using Apps.GoogleSheets.Polling.Models;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common.Polling;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;

namespace Apps.GoogleSheets.Polling
{
    [PollingEventList]
    public class GoogleSheetsPollingList(InvocationContext invocationContext) : BaseInvocable(invocationContext)
    {
        [PollingEvent("On new rows added", "Triggered when new rows are added to the sheet")]
        public async Task<PollingEventResponse<NewRowAddedMemory, IEnumerable<NewRowResult>>> OnNewRowsAdded(
             PollingEventRequest<NewRowAddedMemory> request,
             [PollingEventParameter] SpreadsheetFileRequest spreadsheetFileRequest,
             [PollingEventParameter] SheetRequest sheetRequest)
        {
            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

            var valuesRequest = client.Spreadsheets.Values.Get(spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName);
            var valuesResponse = await valuesRequest.ExecuteAsync();

            int currentRowCount = (valuesResponse.Values != null) ? valuesResponse.Values.Count : 0;
            if (request.Memory == null)
            {
                request.Memory = new NewRowAddedMemory
                {
                    LastRowCount = currentRowCount,
                    LastPollingTime = DateTime.UtcNow,
                    Triggered = false
                };

                return new PollingEventResponse<NewRowAddedMemory, IEnumerable<NewRowResult>>
                {
                    FlyBird = false,
                    Memory = request.Memory,
                    Result = null
                };
            }

            var memory = request.Memory;
            var newRowsList = new List<NewRow>();

            if (valuesResponse.Values != null && currentRowCount > memory.LastRowCount)
            {
                for (int i = memory.LastRowCount; i < currentRowCount; i++)
                {
                    var rowValues = new List<string>();
                    foreach (var cell in valuesResponse.Values[i])
                    {
                        rowValues.Add(cell != null ? cell.ToString() : string.Empty);
                    }

                    var newRow = new NewRow
                    {
                        RowIndex = i + 1,
                        RowValues = rowValues
                    };
                    newRowsList.Add(newRow);
                }
            }

            memory.LastRowCount = currentRowCount;
            memory.LastPollingTime = DateTime.UtcNow;
            memory.Triggered = newRowsList.Any();

            var result = new List<NewRowResult>();
            if (newRowsList.Any())
            {
                var newRowResult = new NewRowResult
                {
                    NewRows = newRowsList
                };
                result.Add(newRowResult);
            }

            var response = new PollingEventResponse<NewRowAddedMemory, IEnumerable<NewRowResult>>
            {
                FlyBird = newRowsList.Any(),
                Memory = memory,
                Result = result
            };
            return response;
        }
    }
}
