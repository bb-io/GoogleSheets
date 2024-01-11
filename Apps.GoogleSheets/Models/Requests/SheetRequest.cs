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
    public class SheetRequest
    {
        [Display("Sheet")]
        [DataSource(typeof(SheetDataSourceHandler))]
        public string SheetName { get; set; }
    }
}
