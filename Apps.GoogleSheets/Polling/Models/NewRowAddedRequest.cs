using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Polling;

namespace Apps.GoogleSheets.Polling.Models
{
    public class NewRowAddedRequest
    {
        [Display("Spreadsheet ID")]
        public string SpreadsheetId { get; set; }

        [Display("Sheet name")]
        public string SheetName { get; set; }
    }
}
