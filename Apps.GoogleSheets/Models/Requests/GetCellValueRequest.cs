﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Models.Requests
{
    public class GetCellValueRequest
    {
        public string SpreadSheetId { get; set; }

        public string SheetName { get; set; }

        public int RowId { get; set; }

        public string Column { get; set; }

    }
}
