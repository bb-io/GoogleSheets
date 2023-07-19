using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Apps.GoogleSheets.Models.Requests;
using Apps.GoogleSheets.Models.Responses;
using Apps.GoogleSheets.Dtos;
using Google.Apis.Sheets.v4.Data;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;
using Blackbird.Applications.Sdk.Common.Actions;

namespace Apps.GoogleSheets
{
    [ActionList]
    public class Actions
    {
        #region Actions

        [Action("Get cell value", Description = "Get specific cell value")]
        public async Task<GetCellValueResponse> GetCellValue(
            IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
            [ActionParameter] GetCellValueRequest input)
        {
            var client = new GoogleSheetsClient(authenticationCredentialsProviders);

            var cellAdress = $"{input.Column}{input.RowId}";
            var sheetValues = await GetSheetValues(client,
                input.SpreadSheetId, input.SheetName, cellAdress, cellAdress);

            return new GetCellValueResponse
            {
                CellValue = sheetValues[0][0]?.ToString() ?? string.Empty
            };
        }

        [Action("Get range of cells", Description = "Get range of cells")]
        public async Task<GetRangeCellsResponse> GetRangeOfCells(
            IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
            [ActionParameter] GetRangeCellsRequest input)
        {
            var client = new GoogleSheetsClient(authenticationCredentialsProviders);

            var cellAdressA = $"{input.ColumnA}{input.RowIdA}";
            var cellAdressB = $"{input.ColumnB}{input.RowIdB}";

            var result = await GetSheetValues(client,
                input.SpreadSheetId, input.SheetName, cellAdressA, cellAdressB);

            return new GetRangeCellsResponse
            {
                Rows = result.Select(v => new SheetRowDto(v)).ToList()
            };
        }

        [Action("Set cell value", Description = "Set cell value")]
        public Task SetCellValue(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
            [ActionParameter] SetCellValueRequest input)
        {
            var client = new GoogleSheetsClient(authenticationCredentialsProviders);

            var cellAdress = $"{input.Column}{input.RowId}";
            var range = $"{input.SheetName}!{cellAdress}:{cellAdress}";

            var valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object> { input.Value } }
            };

            var updateRequest = client.Spreadsheets.Values.Update(valueRange, input.SpreadSheetId, range);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;

            return updateRequest.ExecuteAsync();
        }

        [Action("Clear row", Description = "Clear row")]
        public Task ClearRow(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
            [ActionParameter] ClearRowRequest input)
        {
            var client = new GoogleSheetsClient(authenticationCredentialsProviders);

            var range = $"{input.SheetName}!{input.RowId}:{input.RowId}";
            var clearRequest = client.Spreadsheets.Values.Clear(new(), input.SpreadSheetId, range);
            
            return clearRequest.ExecuteAsync();
        }

        [Action("Add new row to table", Description = "Add new row to detected table")]
        public Task AddNewRow(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
            [ActionParameter] AddNewRowRequest input)
        {
            var client = new GoogleSheetsClient(authenticationCredentialsProviders);

            var tableAdress = $"{input.TableStartColumn}{input.TableStartRowId}";
            var range = $"{input.SheetName}!{tableAdress}:{tableAdress}";

            var valueRange = new ValueRange
            {
                Values = new List<IList<object>> { input.Columns.Cast<object>().ToList() }
            };
            
            var appendRequest = client.Spreadsheets.Values.Append(valueRange, input.SpreadSheetId, range);
            appendRequest.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.InsertDataOption = AppendRequest.InsertDataOptionEnum.INSERTROWS;
            
            return appendRequest.ExecuteAsync();
        }

        #endregion

        #region Utils

        private async Task<IList<IList<object?>>> GetSheetValues(
            GoogleSheetsClient client,
            string sheetId, string sheetName, string cellA, string cellB)
        {
            var range = $"{sheetName}!{cellA}:{cellB}";
            var request = client.Spreadsheets.Values.Get(sheetId, range);

            var response = await request.ExecuteAsync();
            return response.Values;
        }

        #endregion
    }
}