using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Models.Requests;
public class OptionalRangeRequest
{
    [Display("Start Cell", Description = "Cell address (e.g. \"A1\", \"B2\", \"C3\")")]
    public string? StartCell { get; set; }

    [Display("End cell", Description = "Cell address (e.g. \"A1\", \"B2\", \"C3\")")]
    public string? EndCell { get; set; }
}
