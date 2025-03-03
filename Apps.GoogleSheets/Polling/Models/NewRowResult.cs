using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Polling.Models
{
    public class NewRowResult
    {
        [Display("Ne rows calues")]
        public IEnumerable<NewRow> NewRows { get; set; }
    }

    public class NewRow
    {
        public int RowIndex { get; set; }
        public IEnumerable<string> RowValues { get; set; }
    }
}
