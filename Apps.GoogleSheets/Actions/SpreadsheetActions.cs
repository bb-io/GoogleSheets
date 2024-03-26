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
using System.Text.RegularExpressions;
using Apps.GoogleSheets.Models;
using Apps.GoogleSheets.Models.Responses;
using Blackbird.Applications.Sdk.Common.Files;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.Sdk.Glossaries.Utils.Converters;
using Blackbird.Applications.Sdk.Glossaries.Utils.Dtos;
using Blackbird.Applications.Sdk.Utils.Extensions.Files;
using RestSharp;

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
            int newRowIndex;
            if (range != null && range?.Rows != null ) { newRowIndex = range.Rows.Count + 1; }
            else { newRowIndex = 1; }
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
        public async Task<SheetDto> CreateSheet(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] CreateWorksheetRequest createWorksheetRequest)
        {
            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
            var response = await client.Spreadsheets.BatchUpdate(
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
                }, spreadsheetFileRequest.SpreadSheetId).ExecuteAsync();
            
            return new(response.Replies[0].AddSheet.Properties);
        }

        [Action("Get sheet used range", Description = "Get used range")]
        public async Task<RowsDto> GetUsedRange(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] SheetRequest sheetRequest)
        {
            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
            var request = client.Spreadsheets.Values.Get(spreadsheetFileRequest.SpreadSheetId, sheetRequest.SheetName);
            var result = await request.ExecuteAsync();
            if (result != null && result?.Values != null)
            { return new RowsDto() { Rows = result?.Values?.Select(x => x.Select(y => y?.ToString() ?? string.Empty).ToList()).ToList() }; }
            else return new RowsDto() { };
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

                if (columnName == $"{Term} ({languageCode})" || columnName == $"{Variations} ({languageCode})" ||
                    columnName == $"{Notes} ({languageCode})")
                    return string.Empty;

                return null;
            }

            await using var glossaryStream = await _fileManagementClient.DownloadAsync(glossary.Glossary);
            var blackbirdGlossary = await glossaryStream.ConvertFromTBX();
            
            var client = new GoogleSheetsClient(InvocationContext.AuthenticationCredentialsProviders);
            var sheetTitle = blackbirdGlossary.Title ?? Path.GetFileNameWithoutExtension(glossary.Glossary.Name)!;

            var spreadsheet = await client.Spreadsheets.Get(spreadsheetFileRequest.SpreadSheetId).ExecuteAsync();
            var sheet = spreadsheet.Sheets.FirstOrDefault(sheet => sheet.Properties.Title == sheetTitle)?.Properties;
            
            if (sheet != null && (overwriteSheet == null || overwriteSheet.Value == false))
                sheetTitle += $" {DateTime.Now.ToString("g")}";

            if (sheet == null || (sheet != null && (overwriteSheet == null || overwriteSheet.Value == false)))
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
            else
            {
                await client.Spreadsheets.Values
                    .Clear(new ClearValuesRequest(), spreadsheetFileRequest.SpreadSheetId, sheetTitle).ExecuteAsync();
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

            var rowsToAdd = new List<IList<object>>();
            rowsToAdd.Add(
                new List<object>(new[] { Id, Definition, SubjectField, Notes }.Concat(languageRelatedColumns)));

            foreach (var entry in blackbirdGlossary.ConceptEntries)
            {
                var languageRelatedValues = (IEnumerable<string>)languagesPresent
                    .SelectMany(languageCode =>
                        languageRelatedColumns
                            .Select(column => GetColumnValue(column, entry, languageCode)))
                    .Where(value => value != null);

                rowsToAdd.Add(new List<object>(new[]
                {
                    string.IsNullOrWhiteSpace(entry.Id) ? Guid.NewGuid().ToString() : entry.Id, 
                    entry.Definition ?? "", 
                    entry.SubjectField ?? "",
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
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
            await updateRequest.ExecuteAsync();

            return new(sheet);
        }

        [Action("Export glossary", Description = "Export glossary from sheet")]
        public async Task<GlossaryWrapper> ExportGlossary(
            [ActionParameter] SpreadsheetFileRequest spreadsheetFileRequest,
            [ActionParameter] SheetRequest sheetRequest,
            [ActionParameter] [Display("Title")] string? title,
            [ActionParameter] [Display("Source description")]
            string? sourceDescription)
        {
            var rows = await GetUsedRange(spreadsheetFileRequest, sheetRequest);
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

            await using var glossaryStream = glossary.ConvertToTBX();
            var glossaryFileReference = 
                await _fileManagementClient.UploadAsync(glossaryStream, MediaTypeNames.Text.Xml, $"{title}.tbx");
            return new() { Glossary = glossaryFileReference };
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

            var response = await request.ExecuteAsync();
            return response.Values;
        }

        #endregion
    }
}