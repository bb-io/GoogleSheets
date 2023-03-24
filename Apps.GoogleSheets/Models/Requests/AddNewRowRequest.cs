using Apps.GoogleSheets.Dtos;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Models.Requests
{
    public class AddNewRowRequest
    {
        public string SpreadSheetId { get; set; }

        public string SheetName { get; set; }

        public int TableStartRowId { get; set; }

        public string TableStartColumn { get; set; }

        public IEnumerable<string> Columns{ get; set; }
}
}
