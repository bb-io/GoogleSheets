using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Requests
{
    public class SpreadsheetDto
    {
        [Display("Spreadsheet ID")]
        public string Id { get; set; }

        [Display("Spreadsheet URL")]
        public string Url { get; set; }

        [Display("Title")]
        public string Title { get; set; }
    }
}
