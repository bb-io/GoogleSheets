using Blackbird.Applications.Sdk.Common.Exceptions;
using System.Text.RegularExpressions;

namespace Apps.GoogleSheets.Extensions
{
    public static class StringExtensions
    {
        public static string Reverse(this string str)
        {
            char[] reversedString = str.ToCharArray();
            Array.Reverse(reversedString);
            return new String(reversedString);
        }

        public static int ToExcelColumnIndex(this string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new PluginMisconfigurationException("Address value is not valid");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }

        public static (int, int) ToExcelColumnAndRow(this string str)
        {
            var regex = @"([A-Z]+)(\d+)";
            var match = Regex.Match(str, regex);
            if (!match.Success)
                throw new PluginMisconfigurationException($"{str} is not a valid cell address");
            var column = match.Groups[1].Value.ToExcelColumnIndex();
            var row = int.Parse(match.Groups[2].Value);
            return (column, row);
        }
    }
}
