using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Polling.Models
{
    public class NewRowAddedMemory
    {
        public DateTime? LastPollingTime { get; set; }

        public bool Triggered { get; set; }

        public int LastRowCount { get; set; }
    }
}
