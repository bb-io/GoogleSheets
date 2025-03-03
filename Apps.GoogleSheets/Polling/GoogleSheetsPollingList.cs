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
            if (request.Memory == null)
            {
                request.Memory = new NewRowAddedMemory
                {
                    LastRowCount = 0,
                    LastPollingTime = DateTime.UtcNow,
                    Triggered = false
                };
            }

            var memory = request.Memory;

            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

            var valuesRequest = client.Spreadsheets.Values.Get(spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName);
            var valuesResponse = await valuesRequest.ExecuteAsync();

            int currentRowCount = 0;
            if (valuesResponse.Values != null)
            {
                currentRowCount = valuesResponse.Values.Count;
            }
            else
            {
                currentRowCount = 0;
            }

            var newRowsList = new List<NewRow>();

            if (valuesResponse.Values != null && currentRowCount > memory.LastRowCount)
            {
                for (int i = memory.LastRowCount; i < currentRowCount; i++)
                {
                    var rowValues = new List<string>();
                    foreach (var cell in valuesResponse.Values[i])
                    {
                        if (cell != null)
                        {
                            rowValues.Add(cell.ToString());
                        }
                        else
                        {
                            rowValues.Add(string.Empty);
                        }
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
            if (newRowsList.Any())
            {
                memory.Triggered = true;
            }
            else
            {
                memory.Triggered = false;
            }

            var result = new List<NewRowResult>();
            if (newRowsList.Any())
            {
                var newRowResult = new NewRowResult();
                newRowResult.NewRows = newRowsList;
                result.Add(newRowResult);
            }

            var response = new PollingEventResponse<NewRowAddedMemory, IEnumerable<NewRowResult>>();
            if (newRowsList.Any())
            {
                response.FlyBird = true;
            }
            else
            {
                response.FlyBird = false;
            }
            response.Memory = memory;
            response.Result = result;
            return response;
        }
    }
}
