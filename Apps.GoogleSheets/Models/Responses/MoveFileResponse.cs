using Blackbird.Applications.Sdk.Common;

namespace Apps.GoogleSheets.Models.Responses
{
    public class MoveFileResponse
    {
        [Display("File ID")]
        public string Id { get; set; } = default!;

        [Display("Name")]
        public string Name { get; set; } = default!;

        [Display("Parents")]
        public IEnumerable<string>? Parents { get; set; }

        [Display("URL")]
        public string? Url { get; set; }
    }
}
