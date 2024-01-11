using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Models.Requests
{
    public class UpdateRowRequest
    {
        public List<string> Row { get; set; }

        [Display("Start cell address", Description = "Cell address (e.g. \"A1\", \"B2\", \"C3\")")]
        public string CellAddress { get; set; }
    }
}
