using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Dtos
{
    public class SheetRowDto
    {
        public IEnumerable<SheetColumnDto> Columns { get; set; }
    }
}
