using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.GoogleSheets.Extensions
{
    public static class IntExtensions
    {
        // https://www.geeksforgeeks.org/find-excel-column-name-given-number/
        public static string ToExcelColumnAddress(this int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int rem = columnNumber % 26;
                if (rem == 0)
                {
                    columnName += "Z";
                    columnNumber = (columnNumber / 26) - 1;
                }
                else
                {
                    columnName += (char)((rem - 1) + 'A');
                    columnNumber = columnNumber / 26;
                }
            }
            return columnName.Reverse();
        }
    }
}
