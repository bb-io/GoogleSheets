using Apps.GoogleSheets.Utils;

namespace Apps.GoogleSheets.Extensions;

public static class ClientExtensions
{
    public static async Task<IList<IList<object?>>> GetSheetValues(
        this GoogleSheetsClient client,
        string sheetId, 
        string sheetName, 
        string cellA, 
        string cellB)
    {
        var range = $"{sheetName}!{cellA}:{cellB}";
        var request = client.Spreadsheets.Values.Get(sheetId, range);
        
        var response = await ErrorHandler.ExecuteWithErrorHandlingAsync(request.ExecuteAsync);
        if (response?.Values == null || response.Values.Count == 0)
            return new List<IList<object?>>(0);
        
        return response.Values;
    }
}