using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using System.IO;
using Apps.GoogleSheets.Models.Requests;
using Apps.GoogleSheets.Models.Responses;
using Apps.GoogleSheets.Dtos;
using Google.Apis.Sheets.v4.Data;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;
using System.Collections.Generic;
using Blackbird.Applications.Sdk.Common.Actions;

namespace Apps.GoogleSheets
{
    [ActionList]
    public class Actions
    {
        [Action("Get cell value", Description = "Get cell value")]
        public GetCellValueResponse GetCellValue(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
           [ActionParameter] GetCellValueRequest input)
        {
            var client = new GoogleSheetsClient(authenticationCredentialsProviders);
            var cellAdress = $"{input.Column}{input.RowId}";
            var cellValue = client.GetSheetValues(input.SpreadSheetId, input.SheetName, cellAdress, cellAdress)[0][0];
            return new GetCellValueResponse()
            {
                CellValue = cellValue != null ? cellValue.ToString() : string.Empty
            };
        }

        [Action("Get range of cells", Description = "Get range of cells")]
        public GetRangeCellsResponse GetRangeOfCells(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
           [ActionParameter] GetRangeCellsRequest input)
        {
            var client = new GoogleSheetsClient(authenticationCredentialsProviders);
            var cellAdressA = $"{input.ColumnA}{input.RowIdA}";
            var cellAdressB = $"{input.ColumnB}{input.RowIdB}";
            var result = client.GetSheetValues(input.SpreadSheetId, input.SheetName, cellAdressA, cellAdressB);
            var values = result.Select(v => new SheetRowDto() { Columns = v.Select(c => new SheetColumnDto() { Value = c.ToString() }).ToList() }).ToList();
            return new GetRangeCellsResponse()
            {
                Rows = values
            };
        }

        [Action("Set cell value", Description = "Set cell value")]
        public void SetCellValue(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
           [ActionParameter] SetCellValueRequest input)
        {
            var client = new GoogleSheetsClient(authenticationCredentialsProviders);
            var cellAdress = $"{input.Column}{input.RowId}";
            var range = $"{input.SheetName}!{cellAdress}:{cellAdress}";
            IList<IList<object>> test = new List<IList<object>>() { new List<object>() { input.Value } };
            var valueRange = new ValueRange
            {
                Values = test
            };
            var updateRequest = client.Spreadsheets.Values.Update(valueRange, input.SpreadSheetId, range);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.Execute();
        }

        [Action("Clear row", Description = "Clear row")]
        public void ClearRow(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
           [ActionParameter] ClearRowRequest input)
        {
            var range = $"{input.SheetName}!{input.RowId}:{input.RowId}";
            var client = new GoogleSheetsClient(authenticationCredentialsProviders);
            var clearRequest = client.Spreadsheets.Values.Clear(new ClearValuesRequest(), input.SpreadSheetId, range);
            clearRequest.Execute();
        }

        [Action("Add new row to table", Description = "Add new row to detected table")]
        public void AddNewRow(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders,
           [ActionParameter] AddNewRowRequest input)
        {
            var tableAdress = $"{input.TableStartColumn}{input.TableStartRowId}";
            var range = $"{input.SheetName}!{tableAdress}:{tableAdress}";
            var client = new GoogleSheetsClient(authenticationCredentialsProviders);

            IList<IList<object>> test = new List<IList<object>>() { input.Columns.Select(s => (object)s).ToList() };
            var valueRange = new ValueRange
            {
                Values = test
            };
            var appendRequest = client.Spreadsheets.Values.Append(valueRange, input.SpreadSheetId, range);
            appendRequest.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.InsertDataOption = AppendRequest.InsertDataOptionEnum.INSERTROWS;
            appendRequest.Execute();
        }        
    }
}
