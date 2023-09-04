using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Models.Requests
{
    public class AddNewRowRequest
    {
        [DataSource(typeof(SpreadsheetsDataHandler))]
        [Display("Spreadsheet")] public string SpreadSheetId { get; set; }
        [Display("Sheet name")] public string SheetName { get; set; }

        [Display("Cell start")]
        public string? CellStart { get; set; }

        [Display("Cell end")]
        public string? CellEnd { get; set; }
        [Display("Values")] public IEnumerable<string> Columns { get; set; }
    }
}
