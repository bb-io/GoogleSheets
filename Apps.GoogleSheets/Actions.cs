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

namespace Apps.GoogleSheets
{
    [ActionList]
    public class Actions
    {
        [Action("Get cell value", Description = "Get cell value")]
        public GetCellValueResponse GetCellValue(AuthenticationCredentialsProvider authenticationCredentialsProvider,
           [ActionParameter] GetCellValueRequest input)
        {
            var cellAdress = $"{input.Column}{input.RowId}";
            var cellValue = GetSheetValues(authenticationCredentialsProvider.Value, input.SpreadSheetId, input.SheetName, cellAdress, cellAdress)[0][0].ToString();
            return new GetCellValueResponse()
            {
                CellValue = cellValue
            };
        }

        [Action("Get range of cells", Description = "Get range of cells")]
        public GetRangeCellsResponse GetRangeOfCells(AuthenticationCredentialsProvider authenticationCredentialsProvider,
           [ActionParameter] GetRangeCellsRequest input)
        {
            var cellAdressA = $"{input.ColumnA}{input.RowIdA}";
            var cellAdressB = $"{input.ColumnB}{input.RowIdB}";
            var result = GetSheetValues(authenticationCredentialsProvider.Value, input.SpreadSheetId, input.SheetName, cellAdressA, cellAdressB);
            var values = result.Select(v => new SheetRowDto() { Columns = v.Select(c => new SheetColumnDto() { Value = c.ToString() }).ToList() }).ToList();
            return new GetRangeCellsResponse()
            {
                Rows = values
            };
        }

        [Action("Set cell value", Description = "Set cell value")]
        public void SetCellValue(AuthenticationCredentialsProvider authenticationCredentialsProvider,
           [ActionParameter] SetCellValueRequest input)
        {
            var cellAdress = $"{input.Column}{input.RowId}";
            var range = $"{input.SheetName}!{cellAdress}:{cellAdress}";

            var client = GetGoogleSheetsClient(authenticationCredentialsProvider.Value);
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
        public void ClearRow(AuthenticationCredentialsProvider authenticationCredentialsProvider,
           [ActionParameter] ClearRowRequest input)
        {
            var range = $"{input.SheetName}!{input.RowId}:{input.RowId}";
            var client = GetGoogleSheetsClient(authenticationCredentialsProvider.Value);
            var clearRequest = client.Spreadsheets.Values.Clear(new ClearValuesRequest(), input.SpreadSheetId, range);
            clearRequest.Execute();
        }

        [Action("Add new row to table", Description = "Add new row to detected table")]
        public void AddNewRow(AuthenticationCredentialsProvider authenticationCredentialsProvider,
           [ActionParameter] AddNewRowRequest input)
        {
            var tableAdress = $"{input.TableStartColumn}{input.TableStartRowId}";
            var range = $"{input.SheetName}!{tableAdress}:{tableAdress}";
            var client = GetGoogleSheetsClient(authenticationCredentialsProvider.Value);

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

        private IList<IList<object>> GetSheetValues(string serviceAccountConfString, string sheetId, string sheetName,
            string cellA, string cellB)
        {
            var range = $"{sheetName}!{cellA}:{cellB}";
            var request = GetGoogleSheetsClient(serviceAccountConfString).Spreadsheets.Values.Get(sheetId, range);
            var response = request.Execute();
            var values = response.Values;
            return values;
        }

        private SheetsService GetGoogleSheetsClient(string serviceAccountConfString)
        {
            string[] scopes = { SheetsService.Scope.Spreadsheets };
            ServiceAccountCredential? credential = GoogleCredential.FromJson(serviceAccountConfString)
                                                  .CreateScoped(scopes)
                                                  .UnderlyingCredential as ServiceAccountCredential;
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Blackbird"
            });
            return service;
        }   
    }
}
