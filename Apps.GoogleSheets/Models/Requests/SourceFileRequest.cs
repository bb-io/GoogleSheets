using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Files;

namespace Apps.GoogleSheets.Models.Requests
{
    public class SourceFileRequest
    {
        [Display("Source file")]
        public FileReference File {get; set;}

        [Display("Source sheet number")]
        public string Sheet { get; set;}

        [Display("Top left cell", Description = "Using a letter to describe the column and a number for the row index. Examples: A1 or C5")]
        public string? TopLeftCell { get; set; }
    }
}
 