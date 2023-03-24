using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Models.Requests
{
    public class GetRangeCellsRequest
    {
        public string SpreadSheetId { get; set; }

        public string SheetName { get; set; }


        public int RowIdA { get; set; }

        public string ColumnA { get; set; }


        public int RowIdB { get; set; }

        public string ColumnB { get; set; }
    }
}
