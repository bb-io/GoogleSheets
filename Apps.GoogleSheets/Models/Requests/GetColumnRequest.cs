using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Models.Requests
{
    public class GetColumnRequest
    {
        [Display("Column", Description = "Column address (e.g. \"A\", \"B\", \"C\")")]
        public string Column { get; set; }

        [Display("Start Row", Description = "Row number (e.g. \"1\", \"2\", \"3\")")]
        public int StartRow { get; set; }

        [Display("End Row", Description = "Row number (e.g. \"1\", \"2\", \"3\")")]
        public int EndRow { get; set; }
    }
}
