﻿using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.GoogleSheets.Models.Requests
{
    public class ClearRowRequest
    {
        [DataSource(typeof(SpreadsheetsDataHandler))]
        [Display("Spreadsheet")] public string SpreadSheetId { get; set; }
        [Display("Sheet name")] public string SheetName { get; set; }
        [Display("Row")] public string RowId { get; set; }
    }
}