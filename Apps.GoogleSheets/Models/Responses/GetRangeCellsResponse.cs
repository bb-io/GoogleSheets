using Apps.GoogleSheets.Dtos;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Models.Responses
{
    public class GetRangeCellsResponse
    {
        public IEnumerable<SheetRowDto> Rows { get; set; }
    }
}
