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
    public class SpreadsheetFileRequest
    {
        [Display("Spreadsheet")]
        [DataSource(typeof(SpreadsheetFileDataSourceHandler))]
        public string SpreadSheetId { get; set; }
    }
}
