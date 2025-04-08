﻿using Apps.GoogleSheets.DataSourceHandler;
using Blackbird.Applications.Sdk.Common.Dynamic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tests.GoogleSheets.Base;

namespace Tests.GoogleSheets
{
    [TestClass]
    public class DatahandlerTests :TestBase
    {
        [TestMethod]
        public async Task SpreadsheetDatahandler_IsSuccess()
        {
            var handler = new SpreadsheetFileDataSourceHandler(InvocationContext);

            var result = handler.GetData(new DataSourceContext { });

            Console.WriteLine($"Total: {result.Count()}");
            foreach (var item in result)
            {
                Console.WriteLine($"{item.Value}: {item.Key}");
            }

            Assert.IsTrue(result.Count() > 0);
        }
    }
}
