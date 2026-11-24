using Apps.GoogleSheets.Extensions;
using FluentAssertions;

namespace Tests.GoogleSheets;

[TestClass]
public class CellValueExtensionsTests
{
    [TestMethod]
    public void ReplaceEmptyCellPlaceholders_RowWithPlaceholders_ReplacesThemWithEmptyValues()
    {
        var row = new List<string> { "Job 1", "<EMPTY>", " <empty> ", "Done" };

        var result = row.ReplaceEmptyCellPlaceholders();

        result.Should().Equal("Job 1", string.Empty, string.Empty, "Done");
    }

    [TestMethod]
    public void ReplaceEmptyCellPlaceholders_RowWithoutPlaceholders_KeepsValuesUnchanged()
    {
        var row = new List<string> { "-", "EMPTY", "<EMPTY> value", "" };

        var result = row.ReplaceEmptyCellPlaceholders();

        result.Should().Equal("-", "EMPTY", "<EMPTY> value", "");
    }

    [TestMethod]
    public void IsEmptyCell_PlaceholderOrBlankValue_ReturnsTrue()
    {
        new[] { "<EMPTY>", "<Empty>", " ", "", null }
            .Should().OnlyContain(value => value.IsEmptyCell());
    }

    [TestMethod]
    public void IsEmptyCell_RegularValue_ReturnsFalse()
    {
        "Job 1".IsEmptyCell().Should().BeFalse();
    }
}
