﻿using Blackbird.Applications.Sdk.Common;
using Apps.GoogleSheets.Models.Requests;
using Apps.GoogleSheets.Dtos;
using Google.Apis.Sheets.v4.Data;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;
using Blackbird.Applications.Sdk.Common.Actions;
using Blackbird.Applications.Sdk.Common.Invocation;
using Apps.GoogleSheets.Extensions;
using System.Net.Mime;
using System.Text;
using Apps.GoogleSheets.Models.Responses;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;

namespace Apps.GoogleSheets.Actions
{
    [ActionList]
    public class SpreadsheetActions : BaseInvocable
    {
        private readonly IFileManagementClient _fileManagementClient;
        public SpreadsheetActions(InvocationContext invocationContext, IFileManagementClient fileManagementClient) : base(invocationContext)
        {
            _fileManagementClient = fileManagementClient;
        }
        #region Actions

        [Action("Get sheet cell", Description = "Get cell by address")]
        public async Task<CellDto> GetCell(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] SheetRequest sheetRequest,
            [ActionParameter] GetCellRequest input)
        {
            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
            var sheetValues = await GetSheetValues(client,
                spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, input.Cell, input.Cell);
            return new CellDto { Value = sheetValues[0][0]?.ToString() ?? string.Empty };
        }

        [Action("Update sheet cell", Description = "Update cell by address")]
        public async Task<CellDto> UpdateCell(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] SheetRequest sheetRequest,
            [ActionParameter] GetCellRequest cellRequest,
            [ActionParameter] UpdateCellRequest input)
        {
            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
            var range = $"{sheetRequest.SheetName}!{cellRequest.Cell}";

            var valueRange = new ValueRange { Values = new List<IList<object>> { new List<object> { input.Value } } };
            var updateRequest = client.Spreadsheets.Values.Update(valueRange, spreadsheetFileRequest.SpreadSheetId, range);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.IncludeValuesInResponse = true;
            return new CellDto { Value = (await updateRequest.ExecuteAsync()).UpdatedData.Values[0][0].ToString() };
        }

        [Action("Get sheet row", Description = "Get sheet row by address")]
        public async Task<RowDto> GetRow(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] SheetRequest sheetRequest,
            [ActionParameter] GetRowRequest input)
        {
            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

            var result = await GetSheetValues(client,
                spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, $"{input.Column1}{input.RowIndex}", $"{input.Column2}{input.RowIndex}");

            return new RowDto { Row = result.First().Select(x => x?.ToString() ?? string.Empty).ToList() };
        }

        [Action("Add new sheet row", Description = "Adds a new row to the first empty line of the sheet")]
        public async Task<RowDto> AddRow(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] SheetRequest sheetRequest,
            [ActionParameter] InsertRowRequest insertRowRequest)
        {
            var range = await GetUsedRange(spreadsheetFileRequest, sheetRequest);
            var newRowIndex = range.Rows.First().All(x => string.IsNullOrWhiteSpace(x)) ? 1 : range.Rows.Count + 1;
            var startColumn = insertRowRequest.ColumnAddress ?? "A";
            return await UpdateRow(spreadsheetFileRequest, sheetRequest, new UpdateRowRequest { Row = insertRowRequest.Row, CellAddress = startColumn + newRowIndex });
        }

        [Action("Update sheet row", Description = "Update row by start address")]
        public async Task<RowDto> UpdateRow(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] SheetRequest sheetRequest,
            [ActionParameter] UpdateRowRequest updateRowRequest)
        {
            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
            var (startColumn, row) = updateRowRequest.CellAddress.ToExcelColumnAndRow();
            var endColumn = startColumn + updateRowRequest.Row.Count - 1;
            var range = $"{sheetRequest.SheetName}!{startColumn.ToExcelColumnAddress()}{row}:{endColumn.ToExcelColumnAddress()}{row}";
            var valueRange = new ValueRange { Values = new List<IList<object>> { updateRowRequest.Row.Select(x => (object)x).ToList() } };
            var updateRequest = client.Spreadsheets.Values.Update(valueRange, spreadsheetFileRequest.SpreadSheetId, range);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.IncludeValuesInResponse = true;
            var result = await updateRequest.ExecuteAsync();
            return new RowDto() { Row = result.UpdatedData.Values[0].Select(x => x.ToString()).ToList() };
        }

        [Action("Create sheet", Description = "Create sheet")]
        public async Task CreateSheet(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] CreateWorksheetRequest createWorksheetRequest)
        {
            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
            await client.Spreadsheets.BatchUpdate(
                new BatchUpdateSpreadsheetRequest()
                {
                    Requests = new List<Request>()
                    {
                        new Request()
                        {
                            AddSheet = new AddSheetRequest()
                            {
                                Properties = new SheetProperties(){Title = createWorksheetRequest.Name}
                            }
                        }
                    }
                }, spreadsheetFileRequest.SpreadSheetId).ExecuteAsync();
        }

        [Action("Get sheet used range", Description = "Get used range")]
        public async Task<RowsDto> GetUsedRange(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] SheetRequest sheetRequest)
        {
            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
            var request = client.Spreadsheets.Values.Get(spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName);
            var result = await request.ExecuteAsync();
            return new RowsDto() { Rows = result.Values.Select(x => x.Select(y => y?.ToString() ?? string.Empty).ToList()).ToList() };
        }

        [Action("Download sheet CSV file", Description = "Download CSV file")]
        public async Task<FileResponse> DownloadCSV(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] SheetRequest sheetRequest)
        {
            var rows = await GetUsedRange(spreadsheetFileRequest, sheetRequest);
            var csv = new StringBuilder();
            rows.Rows.ForEach(row =>
            {
                csv.AppendLine(string.Join(",", row));
            });

            using var stream = new MemoryStream(Encoding.ASCII.GetBytes(csv.ToString()));
            var csvFile = await _fileManagementClient.UploadAsync(stream, MediaTypeNames.Text.Csv, $"{sheetRequest.SheetName}.csv");
            return new FileResponse() { File = csvFile };
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