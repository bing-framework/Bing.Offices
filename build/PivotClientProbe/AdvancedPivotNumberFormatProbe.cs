using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace Bing.Offices.Tools.PivotClientProbe;

/// <summary>
/// 对比公共透视数字格式入口的真实保存结果。
/// </summary>
internal static class AdvancedPivotNumberFormatProbe
{
    /// <summary>
    /// 保存数字格式诊断工作簿。
    /// </summary>
    /// <param name="output">新建工作簿路径。</param>
    /// <remarks>
    /// 此入口直接调用引擎公共 API，仅用于诊断，不代替 Provider 验收。
    /// </remarks>
    internal static void Run(string output)
    {
        output = Path.GetFullPath(output);
        if (File.Exists(output) || File.Exists(output + ".probe.json"))
            throw new ArgumentException("Use a new number format diagnostic path.");
        using var workbook = new XLWorkbook();
        var source = workbook.AddWorksheet("Data");
        source.Cell("A1").Value = "Group";
        source.Cell("B1").Value = "Amount";
        source.Cell("A2").Value = "North";
        source.Cell("B2").Value = 12.345;
        var before = new List<object>();
        foreach (var mode in new[] { "Fluent", "Property", "CustomId", "DataStyle" })
        {
            var sheet = workbook.AddWorksheet(mode);
            var table = sheet.PivotTables.Add(mode, sheet.Cell("A1"), source.Range("A1:B2"));
            table.RowLabels.Add("Group");
            var value = table.Values.Add("Amount", "Total");
            switch (mode)
            {
                case "Fluent": value.NumberFormat.SetFormat("0.0"); break;
                case "Property": value.NumberFormat.Format = "0.0"; break;
                case "CustomId":
                    value.NumberFormat.SetNumberFormatId(-1);
                    value.NumberFormat.SetFormat("0.0");
                    break;
                case "DataStyle":
                    value.NumberFormat.SetFormat("0.0");
                    table.RowLabels.Get("Group").StyleFormats.DataValuesFormat
                        .ForValueField(value).Style.NumberFormat.Format = "0.0";
                    break;
            }
            before.Add(new { mode, value.NumberFormat.Format, value.NumberFormat.NumberFormatId });
        }
        workbook.SaveAs(output);
        using var reopened = new XLWorkbook(output);
        var after = reopened.Worksheets.Where(sheet => sheet.Name != "Data")
            .Select(sheet => new
            {
                mode = sheet.Name, format = sheet.PivotTables.Single().Values.Single().NumberFormat.Format,
                id = sheet.PivotTables.Single().Values.Single().NumberFormat.NumberFormatId
            }).ToArray();
        using var archive = ZipFile.OpenRead(output);
        var parts = archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal)
                && (entry.FullName.StartsWith("xl/pivotTables/", StringComparison.Ordinal)
                    || entry.FullName == "xl/styles.xml"))
            .ToDictionary(entry => entry.FullName, entry =>
            {
                using var stream = entry.Open();
                return XDocument.Load(stream).ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
            });
        var report = new { diagnosticOnly = true, before, after, parts };
        File.WriteAllText(output + ".probe.json", JsonSerializer.Serialize(report,
            new JsonSerializerOptions { WriteIndented = true }) + "\n", new UTF8Encoding(false));
        Console.WriteLine("number-format-diagnostic-completed");
    }
}
