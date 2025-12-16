using Apps.GoogleSheets.Dtos;
using Apps.GoogleSheets.Extensions;
using Apps.GoogleSheets.Models;
using Apps.GoogleSheets.Models.Requests;
using Apps.GoogleSheets.Models.Responses;
using Apps.GoogleSheets.Utils;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Actions;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Blackbird.Applications.Sdk.Common.Files;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Glossaries.Utils.Converters;
using Blackbird.Applications.Sdk.Glossaries.Utils.Dtos;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using CsvHelper;
using CsvHelper.Configuration;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Globalization;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;
using SheetProperties = Google.Apis.Sheets.v4.Data.SheetProperties;

namespace Apps.GoogleSheets.Actions;

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
        if (string.IsNullOrWhiteSpace(spreadsheetFileRequest?.SpreadSheetId))
            throw new PluginMisconfigurationException("Spreadsheet ID cannot be empty.");

        if (string.IsNullOrWhiteSpace(sheetRequest?.SheetName))
            throw new PluginMisconfigurationException("Sheet title cannot be empty.");

        if (string.IsNullOrWhiteSpace(input?.Column))
            throw new PluginMisconfigurationException("Column cannot be empty.");

        if (!Regex.IsMatch(input.Column, @"^[A-Za-z]+$"))
            throw new PluginMisconfigurationException("Column must contain only letters, e.g., A, B, AA.");

        var rowNumber = ParseRow(input.Row);
        if (rowNumber <= 0)
            throw new PluginMisconfigurationException("Row must be a positive number (e.g., 1, 2, 3).");

        var a1 = $"{input.Column.ToUpperInvariant()}{rowNumber}";

        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

        var sheetValues = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () =>
            await GetSheetValues(client, spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, a1, a1));

        if (sheetValues is null || sheetValues.Count == 0)
            return new CellDto(string.Empty);

        var row = sheetValues[0];
        if (row is null || row.Count == 0)
            return new CellDto(string.Empty);

        return new CellDto(row[0]?.ToString() ?? string.Empty);
    }

    [Action("Update sheet cell", Description = "Update cell by address")]
    public async Task<CellDto> UpdateCell(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest,
        [ActionParameter] GetCellRequest cellRequest,
        [ActionParameter] UpdateCellRequest input)
    {
        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

        await ExpandRowLimits(ParseRow(cellRequest.Row), spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, client);

        var range = $"{sheetRequest.SheetName}!{cellRequest.Column}{ParseRow(cellRequest.Row)}";

        var valueRange = new ValueRange { Values = new List<IList<object>> { new List<object> { input.Value } } };
        var updateRequest = client.Spreadsheets.Values.Update(valueRange, spreadsheetFileRequest.SpreadSheetId, range);
        updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
        updateRequest.IncludeValuesInResponse = true;
        
        var result = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => await updateRequest.ExecuteAsync());
        if (result?.UpdatedData?.Values == null || result.UpdatedData.Values.Count == 0 || result.UpdatedData.Values[0].Count == 0)
        {
            throw new PluginApplicationException("No updated data was returned from the API. Please check your input and try again");
        }
        return new CellDto(result?.UpdatedData.Values[0][0].ToString() ?? string.Empty);
    }

    [Action("Debug", Description = "Can be used only for debugging purposes.")]
    public List<AuthenticationCredentialsProvider> GetAuthenticationCredentialsProviders()
    {
        return InvocationContext.AuthenticationCredentialsProviders.ToList();
    }

    [Action("Get sheet row", Description = "Get sheet row by address")]
    public async Task<RowDto> GetRow(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest,
        [ActionParameter] GetRowRequest input)
    {
        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

        var result = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => await GetSheetValues(client,
            spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, $"{input.Column1}{ParseRow(input.RowIndex)}", $"{input.Column2}{ParseRow(input.RowIndex)}"));
        if (result is null)
        { return new RowDto { Row = new List<string>() }; }

        return new RowDto { Row = result.FirstOrDefault()?.Select(x => x?.ToString() ?? string.Empty)?.ToList() ?? new List<string>() };
    }

    [Action("Add new sheet row", Description = "Adds a new row to the first empty line of the sheet")]
    public async Task<RowDto> AddRow(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest,
        [ActionParameter] InsertRowRequest insertRowRequest)
    {
        if (insertRowRequest?.Row == null || !insertRowRequest.Row.Any())
        {
            throw new PluginMisconfigurationException("The row cannot be null or empty. Please check your input and try again");
        }

        if (insertRowRequest.Row.All(cell => string.IsNullOrWhiteSpace(cell)))
        {
            throw new PluginMisconfigurationException("The row cannot contain only empty values. Please check your input and try again");
        }

        var range = await ErrorHandler.ExecuteWithErrorHandlingAsync(() => GetUsedRange(spreadsheetFileRequest, sheetRequest));
        int newRowIndex;
        if (range != null && range?.Rows != null) { newRowIndex = range.Rows.Count + 1; }
        else { newRowIndex = 1; }
        var startColumn = insertRowRequest.ColumnAddress ?? "A";
        return await ErrorHandler.ExecuteWithErrorHandlingAsync(() => UpdateRow(spreadsheetFileRequest, sheetRequest, new UpdateRowRequest { Row = insertRowRequest.Row, CellAddress = startColumn + newRowIndex }));
    }

    [Action("Clear sheet range", Description = "Clears the values of all cells within a specified range.")]
    public async Task ClearRange(
    [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
    [ActionParameter] SheetRequest sheetRequest,
    [ActionParameter] ClearRangeRequest input)
    {
        if (string.IsNullOrWhiteSpace(spreadsheetFileRequest?.SpreadSheetId))
            throw new PluginMisconfigurationException("Spreadsheet ID cannot be empty.");

        if (string.IsNullOrWhiteSpace(sheetRequest?.SheetName))
            throw new PluginMisconfigurationException("Sheet title cannot be empty.");

        if (string.IsNullOrWhiteSpace(input?.Range))
            throw new PluginMisconfigurationException("Range cannot be empty.");

        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

        await ErrorHandler.ExecuteWithErrorHandlingAsync(async () =>
            await client.Spreadsheets.Values.Clear(
                new ClearValuesRequest(),
                spreadsheetFileRequest.SpreadSheetId,
                $"{sheetRequest.SheetName}!{input.Range}"
            ).ExecuteAsync());
    }

    [Action("Update sheet row", Description = "Update row by start address")]
    public async Task<RowDto> UpdateRow(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest,
        [ActionParameter] UpdateRowRequest updateRowRequest)
    {
        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
        var (startColumn, row) = updateRowRequest.CellAddress.ToExcelColumnAndRow();

        await ExpandRowLimits(row, spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, client);

        var endColumn = startColumn + updateRowRequest.Row.Count - 1;
        var range = $"{sheetRequest.SheetName}!{startColumn.ToExcelColumnAddress()}{row}:{endColumn.ToExcelColumnAddress()}{row}";
        var valueRange = new ValueRange { Values = new List<IList<object>> { updateRowRequest.Row.Select(x => (object)x).ToList() } };
        var updateRequest = client.Spreadsheets.Values.Update(valueRange, spreadsheetFileRequest.SpreadSheetId, range);
        updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
        updateRequest.IncludeValuesInResponse = true;
        var result = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => await updateRequest.ExecuteAsync());
        return new RowDto() { Row = result.UpdatedData.Values[0].Select(x => x.ToString()).ToList() };
    }

    [Action("Create sheet", Description = "Create sheet")]
    public async Task<SheetDto> CreateSheet(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] CreateWorksheetRequest createWorksheetRequest)
    {
        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
        var response = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => await client.Spreadsheets.BatchUpdate(
            new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>
                {
                    new()
                    {
                        AddSheet = new AddSheetRequest
                        {
                            Properties = new SheetProperties {Title = createWorksheetRequest.Name}
                        }
                    }
                }
            }, spreadsheetFileRequest.SpreadSheetId).ExecuteAsync());

        return new(response.Replies[0].AddSheet.Properties);
    }

    [Action("Move file", Description = "Move a Google Drive file to a new parent folder.")]
    public async Task<MoveFileResponse> MoveFile([ActionParameter] MoveFileRequest input)
    {
        if (string.IsNullOrWhiteSpace(input?.FileId))
            throw new PluginMisconfigurationException("File ID cannot be empty. Please check your input and try again.");

        if (string.IsNullOrWhiteSpace(input.NewParentFolderId))
            throw new PluginMisconfigurationException("New parent folder ID cannot be empty. Please check your input and try again.");

        var drive = new GoogleDriveClient(InvocationContext.AuthenticationCredentialsProviders);

        var get = drive.Files.Get(input.FileId);
        get.Fields = "id,name,parents,webViewLink";
        get.SupportsAllDrives = true;

        var meta = await ErrorHandler.ExecuteWithErrorHandlingAsync(get.ExecuteAsync);

        var previousParents = string.Join(",", meta.Parents ?? new List<string>());

        var update = drive.Files.Update(new Google.Apis.Drive.v3.Data.File(), input.FileId);
        update.AddParents = input.NewParentFolderId;
        update.RemoveParents = previousParents;
        update.Fields = "id,name,parents,webViewLink";
        update.SupportsAllDrives = true;

        var updated = await ErrorHandler.ExecuteWithErrorHandlingAsync(update.ExecuteAsync);

        return new MoveFileResponse
        {
            Id = updated.Id,
            Name = updated.Name,
            Parents = updated.Parents,
            Url = updated.WebViewLink
        };
    }

    [Action("Create spreadsheet", Description = "Create a new spreadsheet")]
    public async Task<SpreadsheetDto> CreateSpreadsheet([ActionParameter] CreateSpreadsheetRequest input)
    {
        if (string.IsNullOrWhiteSpace(input.Title))
            throw new PluginMisconfigurationException("Title cannot be empty. Please provide a spreadsheet title and try again");

        var sheets = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

        var spreadsheet = new Spreadsheet
        {
            Properties = new SpreadsheetProperties { Title = input.Title }
        };

        if (!string.IsNullOrWhiteSpace(input.InitialSheetName))
        {
            spreadsheet.Sheets = new List<Sheet>
            {
                new()
                {
                    Properties = new SheetProperties { Title = input.InitialSheetName }
                }
            };
        }

        var created = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () =>
            await sheets.Spreadsheets.Create(spreadsheet).ExecuteAsync());

        if (!string.IsNullOrEmpty(input.FolderId))
        {
            await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => await MoveSpreadsheet(created, input.FolderId));
        }

        return new SpreadsheetDto
        {
            Id = created.SpreadsheetId,
            Url = created.SpreadsheetUrl,
            Title = created.Properties?.Title ?? input.Title
        };
    }

    [Action("Get sheet used range", Description = "Get used range")]
    public async Task<RowsDto> GetUsedRange(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest)
    {
        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
        var request = client.Spreadsheets.Values.Get(spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName);
        var result = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => await request.ExecuteAsync());
        if (result != null && result?.Values != null)
        {
            var rangeIDs = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => GetIdsRange(1, result.Values.Count));
            var rows = result?.Values?.Select(x => x.Select(y => y?.ToString() ?? string.Empty).ToList()).ToList();
            return new RowsDto()
            {
                Rows = rangeIDs.Zip(rows, (id, rowvalues) => new _row { RowId = id, Values = rowvalues }).ToList(),
                RowsCount = (double)result?.Values?.Count
            };
        }
        else return new RowsDto() { };
    }

    [Action("Get range", Description = "Get specific range")]
    public async Task<RowsDto> GetRange(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest,
        [ActionParameter] RangeRequest rangeRequest)
    {
        if (string.IsNullOrWhiteSpace(spreadsheetFileRequest.SpreadSheetId))
            throw new PluginMisconfigurationException("Spreadsheet ID can not be null or empty. Please check your input and try again");
        if (string.IsNullOrWhiteSpace(sheetRequest.SheetName))
            throw new PluginMisconfigurationException("Spreadsheet name can not be null or empty. Please check your input and try again");
        if (string.IsNullOrWhiteSpace(rangeRequest.StartCell))
            throw new PluginMisconfigurationException("Start cell can not be null or empty. Please check your input and try again");
        if (string.IsNullOrWhiteSpace(rangeRequest.EndCell))
            throw new PluginMisconfigurationException("End cell can not be null or empty. Please check your input and try again");


        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
        var result = await GetSheetValues(client,
            spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, rangeRequest.StartCell, rangeRequest.EndCell);
        if (result == null)
        {
            return new RowsDto { Rows = new List<_row>(), RowsCount = 0 };
        }
        var (startColumn, startRow) = rangeRequest.StartCell.ToExcelColumnAndRow();
        var (endColumn, endRow) = rangeRequest.EndCell.ToExcelColumnAndRow();
        var rangeIDs = GetIdsRange(startRow, endRow);
        var rows = result.Select(x => x.Select(y => y?.ToString() ?? string.Empty).ToList()).ToList();
        return new RowsDto()
        {
            Rows = rangeIDs.Zip(rows, (id, rowvalues) => new _row { RowId = id, Values = rowvalues }).ToList(),
            RowsCount = result.Count,
        };
    }

    [Action("Get column", Description = "Get column values")]
    public async Task<ColumnDto> GetColumn(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest,
        [ActionParameter] GetColumnRequest columnRequest)
    {
        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
        var result = await GetSheetValues(client,
            spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, $"{columnRequest.Column}{ParseRow(columnRequest.StartRow)}", $"{columnRequest.Column}{ParseRow(columnRequest.EndRow)}");
        if (result is null)
        {
            return new ColumnDto { Column = new List<string>() };
        }
        return new ColumnDto() { Column = result.Select(x => x.FirstOrDefault()?.ToString() ?? string.Empty).ToList() };
    }

    [Action("Find sheet column", Description = "Providing a row index and a cell value, return column position (letter) where said value is located")]
    public async Task<ColumnResponse> FindColumn(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest,
        [ActionParameter] FindColumnRequest input)
    {
        if (string.IsNullOrWhiteSpace(spreadsheetFileRequest?.SpreadSheetId))
            throw new PluginMisconfigurationException("Spreadsheet ID cannot be empty.");
        if (string.IsNullOrWhiteSpace(sheetRequest?.SheetName))
            throw new PluginMisconfigurationException("Sheet title cannot be empty.");
        if (input is null)
            throw new PluginMisconfigurationException("Input cannot be null.");
        if (string.IsNullOrWhiteSpace(input.RowId))
            throw new PluginMisconfigurationException("Row must be provided.");

        var rowIndex = ParseRow(input.RowId);
        if (rowIndex <= 0)
            throw new PluginMisconfigurationException("Row must be a positive number (e.g., 1, 2, 3).");

        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

        var a1RowRange = $"{sheetRequest.SheetName}!{rowIndex}:{rowIndex}";
        var request = client.Spreadsheets.Values.Get(spreadsheetFileRequest.SpreadSheetId, a1RowRange);

        var response = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => await request.ExecuteAsync());
        if (response?.Values is null || response.Values.Count == 0)
            return null;

        var values = response.Values[0]
            .Select(c => c?.ToString() ?? string.Empty)
            .ToList();

        var target = input.Value ?? string.Empty;
        var colIndex = values.FindIndex(v => string.Equals(v, target, StringComparison.Ordinal));

        if (colIndex < 0)
            return null;

        var columnLetter = (colIndex + 1).ToExcelColumnAddress();
        return new ColumnResponse { Column = columnLetter };
    }

    [Action("Update sheet column", Description = "Update column by start address")]
    public async Task<ColumnDto> UpdateColumn(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest,
        [ActionParameter] UpdateRowRequest updateRowRequest)
    {
        if (updateRowRequest.Row == null || !updateRowRequest.Row.Any())
        {
            throw new PluginMisconfigurationException("Row data cannot be null or empty. Please check your input and try again");
        }

        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
        var (Column, startRow) = updateRowRequest.CellAddress.ToExcelColumnAndRow();
        var endRow = startRow + updateRowRequest.Row.Count - 1;
        await ExpandRowLimits(endRow, spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, client);
        var range = $"{sheetRequest.SheetName}!{updateRowRequest.CellAddress}:{Column}{endRow}";
        var valueRange = new ValueRange
        {
            Values = new List<IList<object>> { updateRowRequest.Row.Select(x => (object)x).ToList() },
            MajorDimension = "COLUMNS"
        };
        var updateRequest = client.Spreadsheets.Values.Update(valueRange, spreadsheetFileRequest.SpreadSheetId, range);
        updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
        updateRequest.IncludeValuesInResponse = true;
        var result = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => await updateRequest.ExecuteAsync());
        return new ColumnDto() { Column = result.UpdatedData.Values[0].Select(x => x.ToString()).ToList() };
    }

    [Action("Find sheet row", Description = "Providing a column address and a value, return row number where said value is located")]
    public async Task<int?> FindRow(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest,
        [ActionParameter] FindRowRequest input)
    {
        var range = await GetUsedRange(spreadsheetFileRequest, sheetRequest);
        var maxRowIndex = range.Rows.Count;
        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
        var result = await GetSheetValues(client,
            spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, $"{input.Column}1", $"{input.Column}{maxRowIndex}");
        if (result is null) { return null; }
        var columnValues = result.Select(x => x.FirstOrDefault()?.ToString() ?? string.Empty).ToList();
        var index = columnValues.IndexOf(input.Value);
        index = index + 1;
        return index == 0 ? null : index;
    }

    private int ColumnLetterToNumber(string column)
    {
        int sum = 0;
        foreach (char c in column)
        {
            sum = sum * 26 + (c - 'A' + 1);
        }
        return sum;
    }

    [Action("Import CSV (Append)", Description = "Import CSV file into Google Sheets")]
    public async Task<SheetDto> ImportCSVAppend(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest,
        [ActionParameter] FileResponse csvFile,
        [ActionParameter] CsvOptions csvOptions,
        [ActionParameter] string? topLeftField = null)
    {
        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

        await using var csvStream = await _fileManagementClient.DownloadAsync(csvFile.File);
        var rows = new List<List<string>>();
        using (var reader = new StreamReader(csvStream, Encoding.UTF8))
        using (var csv = new CsvReader(reader, CreateConfiguration(csvOptions)))
        {
            while (await csv.ReadAsync())
                rows.Add(csv.Parser.Record.ToList());
        }

        int maxCols = rows.Any() ? rows.Max(r => r.Count) : 0;
        foreach (var row in rows)
            while (row.Count < maxCols)
                row.Add(string.Empty);

        int startRow, startCol;
        if (!string.IsNullOrEmpty(topLeftField))
        {
            var m = Regex.Match(topLeftField, @"^([A-Za-z]+)(\d+)$");
            if (!m.Success)
                throw new FormatException("Invalid format—use e.g. 'B3'.");
            startCol = ColumnLetterToNumber(m.Groups[1].Value);
            startRow = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
        }
        else
        {
            var getReq = client.Spreadsheets.Values
                .Get(spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName);
            var getRes = await getReq.ExecuteAsync();
            int existing = getRes.Values?.Count ?? 0;
            startRow = existing + 1;
            startCol = 1;
        }

        int endRow = startRow + rows.Count - 1;
        int endColumn = startCol + maxCols - 1;
        string range = $"{sheetRequest.SheetName}!"
                      + $"{startCol.ToExcelColumnAddress()}{startRow}"
                      + $":{endColumn.ToExcelColumnAddress()}{endRow}";

        var valueRange = new ValueRange
        {
            Values = rows
                .Select(r => (IList<object>)r.Cast<object>().ToList())
                .ToList()
        };

        var updateReq = client.Spreadsheets.Values
            .Update(valueRange, spreadsheetFileRequest.SpreadSheetId, range);
        updateReq.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;

        await ErrorHandler.ExecuteWithErrorHandlingAsync(async () =>
            await updateReq.ExecuteAsync()
        );

        var spreadsheet = await client.Spreadsheets
            .Get(spreadsheetFileRequest.SpreadSheetId)
            .ExecuteAsync();
        var sheet = spreadsheet.Sheets
            .FirstOrDefault(s => s.Properties.Title == sheetRequest.SheetName)?
            .Properties;
        if (sheet == null)
            throw new PluginApplicationException("Sheet not found after update");

        return new SheetDto(sheet);
    }

    private CsvConfiguration CreateConfiguration(CsvOptions csvOptions)
    {
        var config = new CsvConfiguration(CultureInfo.InvariantCulture);
        if (csvOptions.IgnoreBlankLines.HasValue) config.IgnoreBlankLines = csvOptions.IgnoreBlankLines.Value;
        if (csvOptions.NewLine is not null) config.NewLine = csvOptions.NewLine;
        if (csvOptions.Delimiter is not null) config.Delimiter = csvOptions.Delimiter;
        if (csvOptions.Comment is not null && csvOptions.Comment.Length > 1) config.Comment = csvOptions.Comment[0];
        if (csvOptions.Escape is not null && csvOptions.Escape.Length > 1) config.Escape = csvOptions.Escape[0];
        if (csvOptions.Quote is not null && csvOptions.Quote.Length > 1) config.Quote = csvOptions.Quote[0];
        return config;
    }

    [Action("Download sheet as CSV file", Description = "Download CSV file")]
    public async Task<FileResponse> DownloadCSV(
         [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
         [ActionParameter] SheetRequest sheetRequest,
         [ActionParameter] OptionalRangeRequest rangeRequest,
         [ActionParameter] CsvOptions csvOptions
         )
    {
        if (string.IsNullOrWhiteSpace(spreadsheetFileRequest.SpreadSheetId))
            throw new PluginMisconfigurationException("Spreadsheet ID can not be null or empty. Please check your input and try again");
        if (string.IsNullOrWhiteSpace(sheetRequest.SheetName))
            throw new PluginMisconfigurationException("Spreadsheet name can not be null or empty. Please check your input and try again");

        List<List<string>> rows = new List<List<string>>();
        if (!string.IsNullOrWhiteSpace(rangeRequest.StartCell) && !string.IsNullOrWhiteSpace(rangeRequest.EndCell))
        {
            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
            var result = await GetSheetValues(client,
                spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName, rangeRequest.StartCell, rangeRequest.EndCell);
            if (result != null)
            {
                rows = result.Select(r => r?.Select(c => c?.ToString() ?? string.Empty).ToList() ?? new List<string>()).ToList();
            }
        }
        else
        {
            var usedRange = await ErrorHandler.ExecuteWithErrorHandlingAsync(() => GetUsedRange(spreadsheetFileRequest, sheetRequest));
            if (usedRange?.Rows != null)
            {
                rows = usedRange.Rows.Where(r => r != null).Select(r => r!.Values ?? new List<string>()).ToList();
            }
        }

        var columnCount = rows.Any() ? rows.Max(x => x?.Count ?? 0) : 0;
        if (rows.Any())
        {
            foreach (var row in rows)
            {
                var current = row?.Count ?? 0;
                var toAdd = columnCount - current;
                if (toAdd > 0)
                {
                    if (row == null)
                        continue;
                    for (int i = 0; i < toAdd; i++) row.Add(string.Empty);
                }
            }
        }

        using var streamOut = new MemoryStream();
        using (var writer = new StreamWriter(streamOut, leaveOpen: true))
        using (var csv = new CsvWriter(writer, CreateConfiguration(csvOptions)))
        {
            for (int i = 0; i < rows.Count; i++)
            {
                foreach (var field in rows[i])
                {
                    csv.WriteField(field);
                }

                if (i < rows.Count - 1)
                {
                    csv.NextRecord();
                }
            }
        }

        streamOut.Position = 0;
        var csvFile = await _fileManagementClient.UploadAsync(streamOut, MediaTypeNames.Text.Csv, $"{sheetRequest.SheetName}.csv");
        return new FileResponse() { File = csvFile };
    }

    [Action("Download spreadsheet", Description = "Download specific spreadsheet as PDF or XLSX")]
    public async Task<FileResponse> DownloadSpreadsheet(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] FileFormatRequest input)
    {
        if (string.IsNullOrWhiteSpace(spreadsheetFileRequest.SpreadSheetId))
            throw new PluginMisconfigurationException("Spreadsheet ID can not be null or empty. Please check your input and try again");

        var drive = new GoogleDriveClient(InvocationContext.AuthenticationCredentialsProviders);

        async Task<Google.Apis.Drive.v3.Data.File> GetMetaAsync(string id)
        {
            var get = drive.Files.Get(id);
            get.Fields = "id,name,mimeType,webViewLink,shortcutDetails/targetId,shortcutDetails/targetMimeType";
            get.SupportsAllDrives = true;
            return await ErrorHandler.ExecuteWithErrorHandlingAsync(get.ExecuteAsync);
        }

        var meta = await GetMetaAsync(spreadsheetFileRequest.SpreadSheetId);

        if (meta.MimeType == "application/vnd.google-apps.shortcut")
        {
            var targetId = meta.ShortcutDetails?.TargetId;
            if (string.IsNullOrEmpty(targetId))
                throw new PluginMisconfigurationException("The provided file is a shortcut but has no target. Please check the file.");
            meta = await GetMetaAsync(targetId);
        }

        var wantPdf = string.Equals(input.Format, "PDF", StringComparison.OrdinalIgnoreCase);
        var wantXlsx = string.Equals(input.Format, "XLSX", StringComparison.OrdinalIgnoreCase);
        if (!wantPdf && !wantXlsx)
            throw new PluginMisconfigurationException("File format must be PDF or XLSX.");

        if (meta.MimeType == "application/vnd.google-apps.spreadsheet")
        {
            var exportMime = wantPdf
                ? System.Net.Mime.MediaTypeNames.Application.Pdf
                : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            var export = drive.Files.Export(meta.Id, exportMime);
            var stream = await ErrorHandler.ExecuteWithErrorHandlingAsync(export.ExecuteAsStreamAsync);

            var fileName = wantPdf ? $"{meta.Name}.pdf" : $"{meta.Name}.xlsx";
            return new FileResponse
            {
                File = await _fileManagementClient.UploadAsync(stream, exportMime, fileName)
            };
        }

        if (meta.MimeType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            if (wantPdf)
                throw new PluginMisconfigurationException("Can't export an XLSX file to PDF via Drive export. Convert it to Google Sheets first or choose XLSX.");

            var get = drive.Files.Get(meta.Id);
            get.SupportsAllDrives = true;
            get.AcknowledgeAbuse = true;

            using var mem = new MemoryStream();
            await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => await get.DownloadAsync(mem));
            mem.Position = 0;

            return new FileResponse
            {
                File = await _fileManagementClient.UploadAsync(
                    mem,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"{meta.Name}.xlsx")
            };
        }

        throw new PluginMisconfigurationException($"Unsupported file type: {meta.MimeType}. Allowed: Google Sheets or XLSX.");
    }

    [Action("Paste into existing sheet from XLSX file", Description = "Append XLSX spreadsheet content into existing sheet")]
    public async Task PasteFromExcel(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] SheetRequest sheetRequest,
            [ActionParameter] SourceFileRequest xlsxFile)
    {
        var driveClient = new GoogleDriveClient(InvocationContext.AuthenticationCredentialsProviders);
        var sourceStream = await _fileManagementClient.DownloadAsync(xlsxFile.File);

        var fileMetadata = new Google.Apis.Drive.v3.Data.File
        {
            Name = "TempImport",
            MimeType = "application/vnd.google-apps.spreadsheet"
        };

        var request = driveClient.Files.Create(
            fileMetadata,
            sourceStream,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );
        request.Fields = "id";
        var uploadedFile = await request.UploadAsync();
        var tempSpreadsheetId = request.ResponseBody.Id;

        var gsheetClient = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
        var tempSpreadsheet = await gsheetClient.Spreadsheets.Get(tempSpreadsheetId).ExecuteAsync();
        var tempSheet = tempSpreadsheet.Sheets.First();
        var sheetName = tempSheet.Properties.Title;
        var range = $"{sheetName}";
        var response = await gsheetClient.Spreadsheets.Values.Get(tempSpreadsheetId, range).ExecuteAsync();
        int rowCount = response.Values.Count;
        int colCount = response.Values.Max(row => row.Count);
        int tempSheetId = tempSheet.Properties.SheetId.Value;

        int rowIndex = 0;
        int colIndex = 0;

        if (!String.IsNullOrEmpty(xlsxFile.TopLeftCell))
        {
            (colIndex, rowIndex) = xlsxFile.TopLeftCell.ToExcelColumnAndRow();
            rowIndex = rowIndex - 1;
        }

        var targetSpreadsheet = await gsheetClient.Spreadsheets.Get(spreadsheetFileRequest.SpreadSheetId).ExecuteAsync();
        var targetSheetID = targetSpreadsheet.Sheets.First(s => s.Properties.Title == sheetRequest.SheetName)?.Properties.SheetId.Value;

        var copyToRequest = new CopySheetToAnotherSpreadsheetRequest
        {
            DestinationSpreadsheetId = spreadsheetFileRequest.SpreadSheetId
        };

        var copyResponse = await gsheetClient.Spreadsheets.Sheets
            .CopyTo(copyToRequest, tempSpreadsheetId, tempSheetId)
            .ExecuteAsync();

        var copiedSheetId = copyResponse.SheetId.Value;
        await driveClient.Files.Delete(tempSpreadsheetId).ExecuteAsync();

        var copyRequest = new Request
        {
            CopyPaste = new CopyPasteRequest
            {
                Source = new GridRange
                {
                    SheetId = copiedSheetId,
                    StartRowIndex = 0,
                    StartColumnIndex = 0,
                    EndRowIndex = rowCount,
                    EndColumnIndex = colCount
                },
                Destination = new GridRange
                {
                    SheetId = targetSheetID,
                    StartRowIndex = rowIndex,
                    StartColumnIndex = colIndex,
                    EndColumnIndex = colIndex,
                    EndRowIndex = rowIndex

                },
                PasteType = "PASTE_NORMAL"
            }
        };

        var batchUpdate = new BatchUpdateSpreadsheetRequest
        {
            Requests = new List<Request> { copyRequest }
        };

        await gsheetClient.Spreadsheets.BatchUpdate(batchUpdate, spreadsheetFileRequest.SpreadSheetId).ExecuteAsync();
        await gsheetClient.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest
        {
            Requests = new List<Request>{
         new Request { DeleteSheet = new DeleteSheetRequest { SheetId = copiedSheetId}}}
        }, spreadsheetFileRequest.SpreadSheetId).ExecuteAsync();
    }

    [Action("Search spreadsheets", Description = "Search all spreadsheets")]
    public async Task<List<SpreadsheetDto>> SearchSpreadsheets([ActionParameter] GetSpreadsheetsRequest request)
    {
        var driveClient = new GoogleDriveClient(InvocationContext.AuthenticationCredentialsProviders);
        // Search for .gsheet and .xlsx files
        var trashed = request.FetchDeleted?.ToString().ToLower() ?? "false";
        var query = $"(mimeType='application/vnd.google-apps.spreadsheet' " +
                    $"or mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' " +
                    $"or (mimeType='application/vnd.google-apps.shortcut' and " +
                    $"    (shortcutDetails.targetMimeType='application/vnd.google-apps.spreadsheet' " +
                    $"     or shortcutDetails.targetMimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))) " +
                    $"and trashed={trashed}";

        if (!string.IsNullOrEmpty(request.FolderId))
            query += $" and '{request.FolderId}' in parents";

        var spreadsheets = new List<SpreadsheetDto>();
        string? pageToken = null;

        do
        {
            var listReq = driveClient.Files.List();
            listReq.Q = query;
            listReq.Fields =
                "nextPageToken, files(id, name, mimeType, webViewLink, parents, shortcutDetails/targetId, shortcutDetails/targetMimeType)";
            listReq.PageSize = 1000;

            listReq.SupportsAllDrives = true;
            listReq.IncludeItemsFromAllDrives = true;
            listReq.Corpora = "allDrives";
            listReq.Spaces = "drive";

            if (!string.IsNullOrEmpty(pageToken))
                listReq.PageToken = pageToken;

            var res = await ErrorHandler.ExecuteWithErrorHandlingAsync(listReq.ExecuteAsync);

            foreach (var f in res.Files)
            {
                if (f.MimeType == "application/vnd.google-apps.shortcut" &&
                    !string.IsNullOrEmpty(f.ShortcutDetails?.TargetId))
                {
                    var getReq = driveClient.Files.Get(f.ShortcutDetails.TargetId);
                    getReq.Fields = "id, name, mimeType, webViewLink";
                    getReq.SupportsAllDrives = true;

                    var target = await ErrorHandler.ExecuteWithErrorHandlingAsync(getReq.ExecuteAsync);
                    if (target is not null &&
                        (target.MimeType == "application/vnd.google-apps.spreadsheet" ||
                         target.MimeType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    {
                        spreadsheets.Add(new SpreadsheetDto
                        {
                            Id = target.Id,
                            Title = target.Name,
                            Url = target.WebViewLink
                        });
                    }

                    continue;
                }

                if (f.MimeType == "application/vnd.google-apps.spreadsheet" ||
                    f.MimeType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    spreadsheets.Add(new SpreadsheetDto
                    {
                        Id = f.Id,
                        Title = f.Name,
                        Url = f.WebViewLink
                    });
                }
            }

            pageToken = res.NextPageToken;
        } while (!string.IsNullOrEmpty(pageToken));

        return spreadsheets
            .GroupBy(x => x.Id)
            .Select(g => g.First())
            .ToList();
    }

    [Action("Delete sheet", Description = "Delete a sheet within a spreadsheet")]
    public async Task DeleteSheet([ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest, [ActionParameter] SheetRequest sheetRequest)
    {
        if (string.IsNullOrWhiteSpace(spreadsheetFileRequest.SpreadSheetId))
            throw new PluginMisconfigurationException("Spreadsheet ID cannot be empty. Please check your input and try again.");
        if (string.IsNullOrWhiteSpace(sheetRequest.SheetName))
            throw new PluginMisconfigurationException("Sheet name cannot be empty. Please check your input and try again.");

        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);

        var spreadsheet = await ErrorHandler.ExecuteWithErrorHandlingAsync(
            async () => await client.Spreadsheets.Get(spreadsheetFileRequest.SpreadSheetId).ExecuteAsync());

        var targetSheetProps = spreadsheet.Sheets?
            .FirstOrDefault(s => s.Properties?.Title == sheetRequest.SheetName)?.Properties;

        if (targetSheetProps?.SheetId is null)
            throw new PluginMisconfigurationException($"Sheet '{sheetRequest.SheetName}' was not found in the spreadsheet.");

        var sheetCount = spreadsheet.Sheets?.Count ?? 0;
        if (sheetCount <= 1)
            throw new PluginApplicationException("A spreadsheet must contain at least one sheet. Create another sheet before deleting this one.");

        var batchRequest = new BatchUpdateSpreadsheetRequest
        {
            Requests = new List<Request>
        {
            new Request
            {
                DeleteSheet = new DeleteSheetRequest
                {
                    SheetId = targetSheetProps.SheetId
                }
            }
        }
        };

        await ErrorHandler.ExecuteWithErrorHandlingAsync(
            async () => await client.Spreadsheets.BatchUpdate(batchRequest, spreadsheetFileRequest.SpreadSheetId).ExecuteAsync());
    }

    #region Glossaries

    private const string Term = "Term";
    private const string Variations = "Variations";
    private const string Notes = "Notes";
    private const string Id = "ID";
    private const string SubjectField = "Subject field";
    private const string Definition = "Definition";

    [Action("Import glossary", Description = "Import glossary as a sheet")]
    public async Task<SheetDto> ImportGlossary(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] GlossaryWrapper glossary,
        [ActionParameter] [Display("Overwrite existing sheet",
            Description = "Overwrite an existing sheet if it has the same title as the glossary.")]
        bool? overwriteSheet)
    {
        static string? GetColumnValue(string columnName, GlossaryConceptEntry entry, string languageCode)
        {
            var languageSection = entry.LanguageSections.FirstOrDefault(ls => ls.LanguageCode == languageCode);

            if (languageSection != null)
            {
                if (columnName == $"{Term} ({languageCode})")
                    return languageSection.Terms.FirstOrDefault()?.Term;

                if (columnName == $"{Variations} ({languageCode})")
                {
                    var variations = languageSection.Terms.Skip(1).Select(term => term.Term);
                    return string.Join(';', variations);
                }

                if (columnName == $"{Notes} ({languageCode})")
                {
                    var notes = languageSection.Terms.Select(term =>
                        term.Notes == null ? string.Empty : term.Term + ": " + string.Join(';', term.Notes));
                    return string.Join(";; ", notes.Where(note => note != string.Empty));
                }

                return null;
            }

            if (columnName == $"{Term} ({languageCode})" ||
                columnName == $"{Variations} ({languageCode})" ||
                columnName == $"{Notes} ({languageCode})")
            {
                return string.Empty;
            }

            return null;
        }

        await using var originalStream = await _fileManagementClient.DownloadAsync(glossary.Glossary);

        await using var ms = new MemoryStream();
        await originalStream.CopyToAsync(ms);
        ms.Position = 0;

        await ValidateTbxOrThrow(ms, glossary.Glossary);

        ms.Position = 0;

        var blackbirdGlossary = await ms.ConvertFromTbx();

        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
        var sheetTitle = blackbirdGlossary.Title ?? Path.GetFileNameWithoutExtension(glossary.Glossary.Name)!;

        var spreadsheet = await client.Spreadsheets.Get(spreadsheetFileRequest.SpreadSheetId).ExecuteAsync();
        var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == sheetTitle)?.Properties;

        if (sheet != null && (overwriteSheet == null || overwriteSheet.Value == false))
            sheetTitle += $" {DateTime.Now:g}";

        if (sheet == null || (sheet != null && (overwriteSheet == null || overwriteSheet.Value == false)))
        {
            sheet = (await client.Spreadsheets.BatchUpdate(
                new BatchUpdateSpreadsheetRequest
                {
                    Requests = new List<Request>
                    {
                        new()
                        {
                            AddSheet = new AddSheetRequest
                            {
                                Properties = new SheetProperties { Title = sheetTitle }
                            }
                        }
                    }
                }, spreadsheetFileRequest.SpreadSheetId).ExecuteAsync()).Replies[0].AddSheet.Properties;
        }
        else
        {
            await client.Spreadsheets.Values
                .Clear(new ClearValuesRequest(), spreadsheetFileRequest.SpreadSheetId, sheetTitle)
                .ExecuteAsync();
        }

        var languagesPresent = blackbirdGlossary.ConceptEntries
            .SelectMany(entry => entry.LanguageSections)
            .Select(section => section.LanguageCode)
            .Distinct()
            .ToList();

        var languageRelatedColumns = languagesPresent
            .SelectMany(language => new[] { Term, Variations, Notes }
                .Select(suffix => $"{suffix} ({language})"))
            .ToList();

        var rowsToAdd = new List<IList<object>>
        {
            new List<object>(new[] { Id, Definition, SubjectField, Notes }.Concat(languageRelatedColumns))
        };

        foreach (var entry in blackbirdGlossary.ConceptEntries)
        {
            var languageRelatedValues = (IEnumerable<string>)languagesPresent
                .SelectMany(languageCode =>
                    languageRelatedColumns
                        .Select(column => GetColumnValue(column, entry, languageCode)))
                .Where(value => value != null)!;

            rowsToAdd.Add(new List<object>(new[]
            {
                string.IsNullOrWhiteSpace(entry.Id) ? Guid.NewGuid().ToString() : entry.Id,
                entry.Definition ?? string.Empty,
                entry.SubjectField ?? string.Empty,
                string.Join(';', entry.Notes ?? Enumerable.Empty<string>())
            }.Concat(languageRelatedValues)));
        }

        var startColumn = 1;
        var startRow = 1;
        var endColumn = startColumn + rowsToAdd[0].Count - 1;

        var range =
            $"{sheet.Title}!{startColumn.ToExcelColumnAddress()}{startRow}:{endColumn.ToExcelColumnAddress()}{rowsToAdd.Count}";
        var valueRange = new ValueRange { Values = rowsToAdd };
        var updateRequest =
            client.Spreadsheets.Values.Update(valueRange, spreadsheetFileRequest.SpreadSheetId, range);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        await updateRequest.ExecuteAsync();

        return new SheetDto(sheet);
    }

    [Action("Export glossary", Description = "Export glossary from sheet")]
    public async Task<GlossaryWrapper> ExportGlossary(
        [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
        [ActionParameter] SheetRequest sheetRequest,
        [ActionParameter][Display("Title")] string? title,
        [ActionParameter] [Display("Source description")]
        string? sourceDescription)
    {
        var rows = await GetUsedRangeForGlossary(spreadsheetFileRequest, sheetRequest);

        if (rows.Rows == null || !rows.Rows.Any())
        {
            throw new PluginApplicationException("The sheet rows are empty. Please check your inputs and try again");
        }

        var maxLength = rows.Rows.Max(list => list.Count);

        var parsedGlossary = new Dictionary<string, List<string>>();

        for (var i = 0; i < maxLength; i++)
        {
            parsedGlossary[rows.Rows[0][i]] = new List<string>(rows.Rows.Skip(1)
                .Select(row => i < row.Count ? row[i] : string.Empty));
        }

        var glossaryConceptEntries = new List<GlossaryConceptEntry>();

        var entriesCount = rows.Rows.Count - 1;

        for (var i = 0; i < entriesCount; i++)
        {
            string entryId = null;
            string? entryDefinition = null;
            string? entrySubjectField = null;
            List<string>? entryNotes = null;

            var languageSections = new List<GlossaryLanguageSection>();

            foreach (var column in parsedGlossary)
            {
                var columnName = column.Key;
                var columnValues = column.Value;

                switch (columnName)
                {
                    case Id:
                        entryId = i < columnValues.Count ? columnValues[i].Trim() : string.Empty;

                        if (string.IsNullOrWhiteSpace(entryId))
                            entryId = Guid.NewGuid().ToString();

                        break;

                    case Definition:
                        entryDefinition = i < columnValues.Count ? columnValues[i].Trim() : string.Empty;

                        if (string.IsNullOrWhiteSpace(entryDefinition))
                            entryDefinition = null;

                        break;

                    case SubjectField:
                        entrySubjectField = i < columnValues.Count ? columnValues[i].Trim() : string.Empty;

                        if (string.IsNullOrWhiteSpace(entrySubjectField))
                            entrySubjectField = null;

                        break;

                    case Notes:
                        entryNotes = (i < columnValues.Count ? columnValues[i] : string.Empty).Split(';')
                            .Select(value => value.Trim()).ToList();

                        if (entryNotes.All(string.IsNullOrWhiteSpace))
                            entryNotes = null;

                        break;

                    case var languageTerm when new Regex($@"{Term} \(.*?\)").IsMatch(languageTerm):
                        var languageCode = new Regex($@"{Term} \((.*?)\)").Match(languageTerm).Groups[1].Value;
                        if (i < columnValues.Count && !string.IsNullOrWhiteSpace(columnValues[i]))
                            languageSections.Add(new(languageCode,
                                new List<GlossaryTermSection>(new GlossaryTermSection[]
                                    { new(columnValues[i].Trim()) })));
                        break;

                    case var termVariations when new Regex($@"{Variations} \(.*?\)").IsMatch(termVariations):
                        if (i < columnValues.Count && !string.IsNullOrWhiteSpace(columnValues[i]))
                        {
                            languageCode = new Regex($@"{Variations} \((.*?)\)").Match(termVariations).Groups[1]
                                .Value;
                            var targetLanguageSectionIndex =
                                languageSections.FindIndex(section => section.LanguageCode == languageCode);

                            var terms = columnValues[i]
                                .Split(';')
                                .Select(term => new GlossaryTermSection(term.Trim()));

                            if (targetLanguageSectionIndex == -1)
                                languageSections.Add(new(languageCode, new List<GlossaryTermSection>(terms)));
                            else
                                languageSections[targetLanguageSectionIndex].Terms.AddRange(terms);
                        }

                        break;

                    case var termNotes when new Regex($@"{Notes} \(.*?\)").IsMatch(termNotes):
                        if (i < columnValues.Count)
                        {
                            languageCode = new Regex($@"{Notes} \((.*?)\)").Match(termNotes).Groups[1].Value;
                            var targetLanguageSectionIndex =
                                languageSections.FindIndex(section => section.LanguageCode == languageCode);

                            var notesDictionary = columnValues[i]
                                .Split(";; ")
                                .Select(note => note.Split(": "))
                                .Where(note => note.Length > 1)
                                .Select(note => new { Term = note[0], Notes = note[1] })
                                .ToDictionary(value => value.Term.Trim(),
                                    value => value.Notes.Split(';').Select(note => note.Trim()));

                            foreach (var termNotesPair in notesDictionary)
                            {
                                var targetTermIndex = languageSections[targetLanguageSectionIndex].Terms
                                    .FindIndex(term => term.Term == termNotesPair.Key);
                                languageSections[targetLanguageSectionIndex].Terms[targetTermIndex].Notes =
                                    termNotesPair.Value.ToList();
                            }
                        }

                        break;
                }
            }

            var entry = new GlossaryConceptEntry(entryId, languageSections)
            {
                Definition = entryDefinition,
                Notes = entryNotes,
                SubjectField = entrySubjectField
            };
            glossaryConceptEntries.Add(entry);
        }

        if (title == null)
            title = sheetRequest.SheetName;

        var glossary = new Glossary(glossaryConceptEntries)
        {
            Title = title,
            SourceDescription = sourceDescription
                                ??
                                $"Glossary export from Google Sheets on {DateTime.Now.ToLocalTime().ToString("F")}"
        };

        await using var glossaryStream = glossary.ConvertToTbx();
        var glossaryFileReference =
            await _fileManagementClient.UploadAsync(glossaryStream, MediaTypeNames.Text.Xml, $"{title}.tbx");
        return new() { Glossary = glossaryFileReference };
    }

    private static async Task ValidateTbxOrThrow(Stream stream, FileReference file)
    {
        if (stream.CanSeek) stream.Seek(0, SeekOrigin.Begin);
        using var ms = new MemoryStream();
        await stream.CopyToAsync(ms);
        ms.Position = 0;

        using var sr = new StreamReader(ms, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
        var buf = new char[4096];
        var n = await sr.ReadAsync(buf, 0, buf.Length);
        var head = new string(buf, 0, Math.Max(0, n)).TrimStart();

        if (string.IsNullOrWhiteSpace(head) || head[0] != '<')
            throw new PluginMisconfigurationException("Invalid file format: expected TBX.");

        if (head.IndexOf("<tbx", StringComparison.OrdinalIgnoreCase) < 0)
            throw new PluginMisconfigurationException("Invalid TBX: missing <tbx> root element.");

        if (head.IndexOf("urn:iso:std:iso:30042", StringComparison.OrdinalIgnoreCase) < 0)
            throw new PluginMisconfigurationException("Invalid TBX: required TBX namespace not found.");

        if (stream.CanSeek) stream.Seek(0, SeekOrigin.Begin);
    }
    private async Task<SimplerRowsDto> GetUsedRangeForGlossary(SpreadsheetFileRequest spreadsheetFileRequest, SheetRequest sheetRequest)
    {
        var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
        var request = client.Spreadsheets.Values.Get(spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName);
        var result = await request.ExecuteAsync();
        if (result != null && result?.Values != null)
        {
            return new SimplerRowsDto() { Rows = result?.Values?.Select(x => x.Select(y => y?.ToString() ?? string.Empty).ToList()).ToList() };
        }
        else return new SimplerRowsDto() { };
    }

    #endregion

    #endregion

    #region Utils

    private async Task<IList<IList<object?>>> GetSheetValues(
        GoogleSheetsClient client,
        string sheetId, string sheetName, string cellA, string cellB)
    {
        var range = $"{sheetName}!{cellA}:{cellB}";
        var request = client.Spreadsheets.Values.Get(sheetId, range);

        var response = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => await request.ExecuteAsync());
        if (response?.Values == null || response.Values.Count == 0)
            return new List<IList<object?>>(0);
        return response?.Values;
    }
    private List<int> GetIdsRange(int start, int end)
    {
        var myList = new List<int>();
        for (var i = start; i <= end; i++)
        {
            myList.Add(i);
        }
        return myList;
    }

    private async Task ExpandRowLimits(int rowNumber, string spreadSheetId, string sheetName,
        GoogleSheetsClient client)
    {
        var spreadSheetRequest = client.Spreadsheets.Get(spreadSheetId);
        var spreadSheet = await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => 
            await spreadSheetRequest.ExecuteAsync()
        );
        
        var sheet = spreadSheet.Sheets.FirstOrDefault(x => x.Properties.Title == sheetName);
        var rowCount = sheet.Properties.GridProperties.RowCount;

        var expandLength = rowNumber - rowCount;

        if (expandLength > 0)
        {
            var expandRequest = client.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        AppendDimension = new()
                        {
                            SheetId = sheet.Properties.SheetId,
                            Dimension = "ROWS",
                            Length = expandLength
                        }
                    }
                }
            }, spreadSheetId);
            await ErrorHandler.ExecuteWithErrorHandlingAsync(async () => await expandRequest.ExecuteAsync());
        }
    }

    private static int ParseRow(string row)
    {
        if (int.TryParse(row, out int value))
        {
            return value;
        }
        throw new PluginMisconfigurationException("The row value should be a number, e.g. 1");
    }

    private async Task MoveSpreadsheet(Spreadsheet spreadsheet, string folderId)
    {
        var drive = new GoogleDriveClient(InvocationContext.AuthenticationCredentialsProviders);
        string spreadsheetId = spreadsheet.SpreadsheetId;

        var getRequest = drive.Files.Get(spreadsheetId);
        getRequest.Fields = "parents";
        var file = await getRequest.ExecuteAsync();

        var previousParents = string.Join(",", file.Parents ?? new List<string>());
        var updateRequest = drive.Files.Update(
            new Google.Apis.Drive.v3.Data.File(),
            spreadsheetId
        );

        updateRequest.AddParents = folderId;
        updateRequest.RemoveParents = previousParents;
        updateRequest.Fields = "id, parents";
        updateRequest.SupportsAllDrives = true;

        await updateRequest.ExecuteAsync();
    }

    #endregion
}