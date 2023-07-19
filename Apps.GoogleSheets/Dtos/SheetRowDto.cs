namespace Apps.GoogleSheets.Dtos
{
    public class SheetRowDto
    {
        public IEnumerable<SheetColumnDto> Columns { get; set; }

        public SheetRowDto(IList<object?> data)
        {
            Columns = data.Select(c => new SheetColumnDto(c?.ToString())).ToList();
        }
    }
}
