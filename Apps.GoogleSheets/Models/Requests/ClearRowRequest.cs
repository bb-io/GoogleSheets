using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Models.Requests
{
    public class ClearRowRequest
    {
        public string SpreadSheetId { get; set; }

        public string SheetName { get; set; }

        public int RowId { get; set; }
    }
}
