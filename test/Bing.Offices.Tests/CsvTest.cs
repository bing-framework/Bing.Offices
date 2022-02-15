using System;
using System.Data;
using Xunit;

namespace Bing.Offices.Tests
{
    public class CsvTest
    {
        [Fact]
        public void Test_ExportByDataTable()
        {
            var dt = new DataTable();
            dt.Columns.AddRange(new[]
            {
                new DataColumn("Name"),
                new DataColumn("Age"),
                new DataColumn("Desc"),
                new DataColumn("Separator"),
                new DataColumn("Quote"),
            });
            for (var i = 0; i < 10; i++)
            {
                var row = dt.NewRow();
                row.ItemArray = new object[] { $"Test_{i}", i + 10, $"Desc_{i}",$"Separator , {i}",$"Quote , \" {i}" };
                dt.Rows.Add(row);
            }

            CsvHelper.ToCsvFile(dt, $"D:\\test_{DateTime.Now:yyyyMMddHHmmss_fff}.csv");
        }
    }
}