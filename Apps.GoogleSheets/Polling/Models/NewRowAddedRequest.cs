using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Polling;

namespace Apps.GoogleSheets.Polling.Models
{
    public class NewRowAddedRequest
    {
        [Display("Spreadsheet ID")]
        [DataSource(typeof(SpreadsheetFileDataSourceHandler))]
        public string SpreadsheetId { get; set; }

        [Display("Sheet name")]
        [DataSource(typeof(SheetDataSourceHandler))]
        public string SheetName { get; set; }
    }
}
