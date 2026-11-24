namespace Apps.GoogleSheets.Extensions;

public static class CellValueExtensions
{
    public const string EmptyCellPlaceholder = "<EMPTY>";

    public static List<string> ReplaceEmptyCellPlaceholders(this IEnumerable<string> cellValues)
    {
        return cellValues.Select(value => value.IsEmptyCellPlaceholder() ? string.Empty : value).ToList();
    }

    public static bool IsEmptyCell(this string? value)
    {
        return string.IsNullOrWhiteSpace(value) || value.IsEmptyCellPlaceholder();
    }

    private static bool IsEmptyCellPlaceholder(this string? value)
    {
        return string.Equals(value?.Trim(), EmptyCellPlaceholder, StringComparison.OrdinalIgnoreCase);
    }
}
