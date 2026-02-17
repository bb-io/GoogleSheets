using Apps.GoogleSheets.Models.Requests;
using Apps.GoogleSheets.Polling.Models;
using Apps.GoogleSheets.Utils;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common.Polling;

namespace Apps.GoogleSheets.Polling;

[PollingEventList]
public class GoogleSheetsPollingList(InvocationContext invocationContext) : BaseInvocable(invocationContext)
{
    [PollingEvent("On new rows added", "Triggered when new rows are added to the sheet")]
    public async Task<PollingEventResponse<NewRowAddedMemory, NewRowResult>> OnNewRowsAdded(
         PollingEventRequest<NewRowAddedMemory> request,
         [PollingEventParameter] SpreadsheetFileRequest spreadsheetFileRequest,
         [PollingEventParameter] SheetRequest sheetRequest)
    {
        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

        var valuesRequest = client.Spreadsheets.Values.Get(
            spreadsheetFileRequest.SpreadSheetId,
            sheetRequest.SheetName);

        var valuesResponse = await ErrorHandler.ExecuteWithErrorHandlingAsync(
            async () => await valuesRequest.ExecuteAsync());

        var currentRowCount = valuesResponse.Values?.Count ?? 0;

        if (request.Memory == null)
        {
            var initMemory = new NewRowAddedMemory
            {
                LastRowCount = currentRowCount,
                LastPollingTime = DateTime.UtcNow,
                Triggered = false
            };

            return new PollingEventResponse<NewRowAddedMemory, NewRowResult>
            {
                FlyBird = false,
                Memory = initMemory,
                Result = null
            };
        }

        var memory = request.Memory;

        if (valuesResponse.Values == null || currentRowCount <= memory.LastRowCount)
        {
            memory.LastPollingTime = DateTime.UtcNow;
            memory.Triggered = false;

            return new PollingEventResponse<NewRowAddedMemory, NewRowResult>
            {
                FlyBird = false,
                Memory = memory,
                Result = null
            };
        }

        var newRows = new List<NewRow>();

        for (int i = memory.LastRowCount; i < currentRowCount; i++)
        {
            var rowValues = new List<string>();

            foreach (var cell in valuesResponse.Values[i])
                rowValues.Add(cell?.ToString() ?? string.Empty);

            newRows.Add(new NewRow
            {
                RowIndex = i + 1,
                RowValues = rowValues
            });
        }

        memory.LastRowCount = currentRowCount;
        memory.LastPollingTime = DateTime.UtcNow;
        memory.Triggered = newRows.Any();

        var result = new NewRowResult { NewRows = newRows };

        return new PollingEventResponse<NewRowAddedMemory, NewRowResult>
        {
            FlyBird = newRows.Any(),
            Memory = memory,
            Result = result
        };
    }
}
