using Apps.GoogleSheets.Dtos;
using Apps.GoogleSheets.Extensions;
using Apps.GoogleSheets.Models.Requests;
using Apps.GoogleSheets.Polling.Models;
using Apps.GoogleSheets.Utils;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common.Polling;

namespace Apps.GoogleSheets.Polling;

[PollingEventList("Cells")]
public class CellPollingList(InvocationContext invocationContext) : GoogleSheetsInvocable(invocationContext)
{
    [PollingEvent("On value added to cell", "Triggered when a value is added to a specific cell")]
    public async Task<PollingEventResponse<CellValueMemory, CellDto>> OnValueAddedToCell(
         PollingEventRequest<CellValueMemory> request,
         [PollingEventParameter] SpreadsheetFileRequest spreadsheetFileRequest,
         [PollingEventParameter] SheetRequest sheetRequest,
         [PollingEventParameter] GetCellRequest cellAddressRequest)
    {
        string cellAddress = cellAddressRequest.ToString();
        var sheetValues = await ErrorHandler.ExecuteWithErrorHandlingAsync(() =>
            Client.GetSheetValues(spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, cellAddress, cellAddress));
        
        var currentValue = sheetValues.FirstOrDefault()?.FirstOrDefault()?.ToString();
        var memory = new CellValueMemory { LastValue = currentValue };
        
        bool wasEmpty = string.IsNullOrWhiteSpace(request.Memory?.LastValue);
        bool isEmpty = string.IsNullOrWhiteSpace(currentValue);
        if (!wasEmpty || isEmpty)
            return new() { FlyBird = false, Memory = memory };
        
        return new()
        {
            FlyBird = true,
            Memory = memory,
            Result = new CellDto(currentValue ?? string.Empty)
        };
    }
}