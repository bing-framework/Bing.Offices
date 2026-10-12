using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Exports.PivotTables;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI 报表筛选字段在真实 XLSX 中保存类型选择和完整缓存。
/// </summary>
[Collection("NpoiNativePivot")]
public sealed class NpoiPivotTypedFilterWriterTest
{
    /// <summary>
    /// 工作簿来源字段的固定顺序。
    /// </summary>
    private static readonly string[] Headers =
        { "Region", "Amount", "Numeric", "When", "Enabled", "Optional", "Status", "Category" };

    /// <summary>
    /// 包含各种可筛选类型、重复项和原始大小写差异的完整来源。
    /// </summary>
    private static readonly object[][] SourceRows =
    {
        new object[] { "North", 10d, 12d, new DateTime(2025, 1, 1), false, null, "Retail", "Food" },
        new object[] { "South", 20d, 15d, new DateTime(2025, 1, 2), true, "filled", "RETAIL", "Tools" },
        new object[] { "North", 30d, 12d, new DateTime(2025, 1, 1), true, string.Empty, "Closed", "Food" },
        new object[] { "West", 40d, 20d, new DateTime(2025, 1, 3), false, "filled", "Other", "Other" }
    };

    /// <summary>
    /// 验证冷启动并发工作簿、多种类型筛选选择及全量缓存写入。
    /// </summary>
    [Fact]
    public async Task ColdConcurrentAndTypedSelections_ShouldSerializeCompleteSourceAndFilters()
    {
        var coldStartRuns = await Task.WhenAll(Enumerable.Range(0, 8)
            .Select(_ => Task.Run(() => Export("all"))));
        foreach (var bytes in coldStartRuns)
            AssertWorkbook("all", bytes);

        foreach (var mode in new[] { "single", "multiple", "all" })
            AssertWorkbook(mode, Export(mode));
    }

    /// <summary>
    /// 创建工作簿并通过来源预检、原生透视写入器及正式序列化重载导出。
    /// </summary>
    /// <param name="mode">筛选选择模式。</param>
    /// <returns>真实 XLSX 包的完整字节。</returns>
    private static byte[] Export(string mode)
    {
        using var workbook = CreateWorkbook();
        var definition = CreateDefinition(mode);
        var source = NpoiPivotSourceValidator.Validate(workbook, definition, CancellationToken.None);
        var records = NpoiPivotTableWriter.Create(workbook, source, CancellationToken.None);
        using var output = new MemoryStream();
        NpoiSheetContentWriter.Write(workbook,
            new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>(),
            new[] { records }, output, CancellationToken.None);
        Assert.True(output.CanWrite);
        return output.ToArray();
    }

    /// <summary>
    /// 断言透视选择与完整业务来源、类型化缓存记录相符。
    /// </summary>
    /// <param name="mode">筛选选择模式。</param>
    /// <param name="bytes">待检查的真实 XLSX 包。</param>
    private static void AssertWorkbook(string mode, byte[] bytes)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var definitionPath = archive.Entries.Single(entry => entry.FullName.Contains("pivotCacheDefinition",
            StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).FullName;
        var definition = ReadXml(archive, definitionPath).Root;
        var fields = definition.Element(mainNs + "cacheFields").Elements(mainNs + "cacheField").ToArray();
        Assert.Equal(Headers, fields.Select(field => (string)field.Attribute("name")));
        Assert.Equal((uint)SourceRows.Length, (uint)definition.Attribute("recordCount"));
        Assert.Equal((uint)Headers.Length, (uint)definition.Element(mainNs + "cacheFields").Attribute("count"));

        var recordsPath = archive.Entries.Single(entry => entry.FullName.Contains("pivotCacheRecords",
            StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).FullName;
        var recordRoot = ReadXml(archive, recordsPath).Root;
        var recordRows = recordRoot.Elements(mainNs + "r").ToArray();
        Assert.Equal(SourceRows.Length, recordRows.Length);
        Assert.Equal((uint)SourceRows.Length, (uint)recordRoot.Attribute("count"));
        for (var row = 0; row < SourceRows.Length; row++)
        {
            var indices = recordRows[row].Elements(mainNs + "x")
                .Select(item => (int)item.Attribute("v")).ToArray();
            Assert.Equal(Headers.Length, indices.Length);
            for (var column = 0; column < Headers.Length; column++)
            {
                var sharedItems = fields[column].Element(mainNs + "sharedItems").Elements().ToArray();
                AssertValue(SourceRows[row][column], sharedItems[indices[column]],
                    $"cache[{row},{column}] {Headers[column]}");
            }
        }

        var pivotPath = archive.Entries.Single(entry => entry.FullName.Contains("pivotTables/",
            StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).FullName;
        var pivot = ReadXml(archive, pivotPath).Root;
        var pageFields = pivot.Element(mainNs + "pageFields").Elements(mainNs + "pageField").ToArray();
        Assert.Equal(new[] { 2, 3, 4, 5, 6, 7 }, pageFields.Select(field => (int)field.Attribute("fld")));
        var pivotFields = pivot.Element(mainNs + "pivotFields").Elements(mainNs + "pivotField").ToArray();
        AssertFieldSelection(2, NumericSelection(mode), fields, pivotFields, pageFields, mainNs);
        AssertFieldSelection(3, DateSelection(mode), fields, pivotFields, pageFields, mainNs);
        AssertFieldSelection(4, BooleanSelection(mode), fields, pivotFields, pageFields, mainNs);
        AssertFieldSelection(5, OptionalSelection(mode), fields, pivotFields, pageFields, mainNs);
        AssertFieldSelection(6, StatusSelection(mode), fields, pivotFields, pageFields, mainNs);
        AssertFieldSelection(7, Array.Empty<object>(), fields, pivotFields, pageFields, mainNs);

        using var reopened = new XSSFWorkbook(new MemoryStream(bytes, writable: false));
        var data = reopened.GetSheet("Data");
        Assert.Equal(Headers.Length, data.GetRow(0).LastCellNum);
        for (var column = 0; column < Headers.Length; column++)
            Assert.Equal(Headers[column], data.GetRow(0).GetCell(column).StringCellValue);
        for (var row = 0; row < SourceRows.Length; row++)
            for (var column = 0; column < Headers.Length; column++)
                AssertWorkbookValue(SourceRows[row][column], ReadWorkbookValue(data.GetRow(row + 1).GetCell(column), reopened),
                    $"sheet[{row},{column}] {Headers[column]}");
    }

    /// <summary>
    /// 验证单字段的真实缓存项目与报表筛选隐藏状态。
    /// </summary>
    /// <param name="fieldIndex">来源字段零基索引。</param>
    /// <param name="selectedValues">期待选择的类型化字面值。</param>
    /// <param name="cacheFields">缓存字段 XML。</param>
    /// <param name="pivotFields">透视字段 XML。</param>
    /// <param name="pageFields">报表筛选字段 XML。</param>
    /// <param name="mainNs">SpreadsheetML 命名空间。</param>
    private static void AssertFieldSelection(int fieldIndex, object[] selectedValues,
        XElement[] cacheFields, XElement[] pivotFields, XElement[] pageFields, XNamespace mainNs)
    {
        var shared = cacheFields[fieldIndex].Element(mainNs + "sharedItems").Elements().ToArray();
        var selectedIndexes = Enumerable.Range(0, shared.Length).Where(index => selectedValues.Any(value =>
            ValuesEqual(value, ReadSharedValue(shared[index])))).ToHashSet();
        var pivotField = pivotFields[fieldIndex];
        var pivotItems = pivotField.Element(mainNs + "items").Elements(mainNs + "item").ToArray();
        Assert.Equal(Enumerable.Range(0, shared.Length), pivotItems.Select(item => (int)item.Attribute("x")));
        Assert.Equal(selectedIndexes.Count > 1, (bool?)pivotField.Attribute("multipleItemSelectionAllowed") == true);
        foreach (var item in pivotItems)
        {
            var index = (int)item.Attribute("x");
            var expectedHidden = selectedIndexes.Count > 1 && !selectedIndexes.Contains(index);
            Assert.Equal(expectedHidden, XmlConvert.ToBoolean((string)item.Attribute("h") ?? "false"));
        }

        var page = pageFields.Single(field => (int)field.Attribute("fld") == fieldIndex);
        var isAll = selectedValues.Length == 0;
        var expectsSingleItem = !isAll && selectedIndexes.Count == 1;
        var pageItem = page.Attribute("item");
        var context = $"field={fieldIndex}; selectedIndexes=[{string.Join(",", selectedIndexes.OrderBy(index => index))}]; "
            + $"pageField={page}; pivotField={pivotField}";
        Assert.True((pageItem != null) == expectsSingleItem, context);
        if (expectsSingleItem)
            Assert.True(selectedIndexes.Single() == (int)pageItem, context);
    }

    /// <summary>
    /// 验证缓存项值及其基础 CLR 类型与完整输入矩阵一致。
    /// </summary>
    /// <param name="expected">原始工作表值。</param>
    /// <param name="item">共享项 XML。</param>
    /// <param name="context">用于失败信息的坐标上下文。</param>
    private static void AssertValue(object expected, XElement item, string context)
    {
        Assert.True(Equals(expected, ReadSharedValue(item)), context);
    }

    /// <summary>
    /// 验证重开工作簿后的单元格值与原始输入一致。
    /// </summary>
    /// <param name="expected">原始工作表值。</param>
    /// <param name="actual">重开后读取的工作表值。</param>
    /// <param name="context">用于失败信息的坐标上下文。</param>
    private static void AssertWorkbookValue(object expected, object actual, string context)
    {
        Assert.True(Equals(expected, actual), context);
    }

    /// <summary>
    /// 读取缓存共享项中的类型化值。
    /// </summary>
    /// <param name="item">共享项 XML。</param>
    /// <returns>缓存项对应的原始 CLR 值。</returns>
    private static object ReadSharedValue(XElement item)
    {
        var value = (string)item.Attribute("v");
        return item.Name.LocalName switch
        {
            "m" => null,
            "s" => DecodeXstring(value ?? string.Empty),
            "n" => double.Parse(value, CultureInfo.InvariantCulture),
            "b" => value == "1",
            "d" => DateTime.Parse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind),
            _ => throw new InvalidDataException($"Unexpected cache item type {item.Name.LocalName}.")
        };
    }

    /// <summary>
    /// 按工作簿单元格类型读取来源值。
    /// </summary>
    /// <param name="cell">待读取的源单元格。</param>
    /// <param name="workbook">包含日期系统设置的工作簿。</param>
    /// <returns>文本、数值、日期、布尔或空值。</returns>
    private static object ReadWorkbookValue(ICell cell, IWorkbook workbook)
    {
        if (cell == null || cell.CellType == CellType.Blank)
            return null;
        if (cell.CellType == CellType.String)
            return cell.StringCellValue;
        if (cell.CellType == CellType.Boolean)
            return cell.BooleanCellValue;
        if (cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
            return DateUtil.GetJavaDate(cell.NumericCellValue, workbook.IsDate1904());
        if (cell.CellType == CellType.Numeric)
            return cell.NumericCellValue;
        throw new InvalidDataException($"Unexpected source cell type {cell.CellType}.");
    }

    /// <summary>
    /// 解码一次 SpreadsheetML 文本属性转义。
    /// </summary>
    /// <param name="value">Xstring 编码值。</param>
    /// <returns>一次解码后的原文。</returns>
    private static string DecodeXstring(string value)
    {
        var result = new System.Text.StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            if (index + 6 < value.Length && value[index] == '_' && value[index + 1] == 'x'
                && value[index + 6] == '_' && Uri.IsHexDigit(value[index + 2])
                && Uri.IsHexDigit(value[index + 3]) && Uri.IsHexDigit(value[index + 4])
                && Uri.IsHexDigit(value[index + 5]))
            {
                result.Append((char)int.Parse(value.Substring(index + 2, 4), NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture));
                index += 6;
            }
            else
                result.Append(value[index]);
        }
        return result.ToString();
    }

    /// <summary>
    /// 比较筛选选项与缓存值的精确类型语义。
    /// </summary>
    /// <param name="left">筛选配置值。</param>
    /// <param name="right">原始缓存值。</param>
    /// <returns>类型和值均匹配时返回 true。</returns>
    private static bool ValuesEqual(object left, object right)
    {
        if (left == null || right == null)
            return left == right;
        if (left is string leftText && right is string rightText)
            return string.Equals(leftText, rightText, StringComparison.OrdinalIgnoreCase);
        if (left is DateTime leftDate && right is DateTime rightDate)
            return leftDate == rightDate;
        if (left is bool leftBoolean && right is bool rightBoolean)
            return leftBoolean == rightBoolean;
        return left is IConvertible && right is double
            && Convert.ToDouble(left, CultureInfo.InvariantCulture) == (double)right;
    }

    /// <summary>
    /// 创建含各种类型和大小写文本状态的真实来源工作簿。
    /// </summary>
    /// <returns>包含完整 Data 来源的 XLSX 工作簿。</returns>
    private static XSSFWorkbook CreateWorkbook()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var dateStyle = workbook.CreateCellStyle();
        dateStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd hh:mm:ss");
        var header = sheet.CreateRow(0);
        for (var column = 0; column < Headers.Length; column++)
            header.CreateCell(column, CellType.String).SetCellValue(Headers[column]);
        for (var rowIndex = 0; rowIndex < SourceRows.Length; rowIndex++)
        {
            var row = sheet.CreateRow(rowIndex + 1);
            for (var column = 0; column < Headers.Length; column++)
            {
                var cell = row.CreateCell(column);
                switch (SourceRows[rowIndex][column])
                {
                    case null: cell.SetCellType(CellType.Blank); break;
                    case string text: cell.SetCellValue(text); break;
                    case bool boolean: cell.SetCellValue(boolean); break;
                    case DateTime date: cell.SetCellValue(date); cell.CellStyle = dateStyle; break;
                    case double number: cell.SetCellValue(number); break;
                    default: throw new ArgumentException("测试来源值类型未覆盖。");
                }
            }
        }
        return workbook;
    }

    /// <summary>
    /// 创建三种类型筛选选择并附加大小写不敏感的双项状态筛选。
    /// </summary>
    /// <param name="mode">single、multiple 或 all。</param>
    /// <returns>包含多字段报表筛选的透视定义。</returns>
    private static ExcelPivotTableDefinition CreateDefinition(string mode)
    {
        var all = mode == "all";
        return new ExcelPivotTableDefinition
        {
            Name = "TypedFilters",
            SourceSheet = "Data",
            SourceRange = new ExcelRangeDefinition
                { StartRow = 0, StartColumn = 0, EndRow = SourceRows.Length, EndColumn = Headers.Length - 1 },
            TargetSheet = "Summary",
            RowFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
            ValueFields = new[] { new ExcelPivotValueDefinition
                { Field = "Amount", Caption = "Total", Aggregation = ExcelPivotAggregation.Sum } },
            Filters = new[]
            {
                Filter("Numeric", all ? Array.Empty<object>() : mode == "single" ? new object[] { 12d } : new object[] { 12d, 15d }),
                Filter("When", all ? Array.Empty<object>() : mode == "single"
                    ? new object[] { new DateTime(2025, 1, 1) }
                    : new object[] { new DateTime(2025, 1, 1), new DateTime(2025, 1, 2) }),
                Filter("Enabled", all ? Array.Empty<object>() : mode == "single"
                    ? new object[] { false } : new object[] { false, true }),
                Filter("Optional", all ? Array.Empty<object>() : mode == "single"
                    ? new object[] { null } : new object[] { null, "filled" }),
                Filter("Status", all ? Array.Empty<object>() : new object[] { "rEtAiL" }),
                Filter("Category", Array.Empty<object>())
            }
        };
    }

    /// <summary>
    /// 创建单字段报表筛选定义。
    /// </summary>
    /// <param name="field">来源表头。</param>
    /// <param name="values">请求的筛选值；空集合表示全选。</param>
    /// <returns>类型化筛选字段。</returns>
    private static ExcelPivotFilterDefinition Filter(string field, IReadOnlyList<object> values) => new()
    {
        Field = field,
        Values = values
    };

    /// <summary>
    /// 获取数值字段的期待选择值。
    /// </summary>
    /// <param name="mode">筛选选择模式。</param>
    /// <returns>对应模式的数值筛选项。</returns>
    private static object[] NumericSelection(string mode) => mode switch
    {
        "single" => new object[] { 12d },
        "multiple" => new object[] { 12d, 15d },
        _ => Array.Empty<object>()
    };

    /// <summary>
    /// 获取日期字段的期待选择值。
    /// </summary>
    /// <param name="mode">筛选选择模式。</param>
    /// <returns>对应模式的日期筛选项。</returns>
    private static object[] DateSelection(string mode) => mode switch
    {
        "single" => new object[] { new DateTime(2025, 1, 1) },
        "multiple" => new object[] { new DateTime(2025, 1, 1), new DateTime(2025, 1, 2) },
        _ => Array.Empty<object>()
    };

    /// <summary>
    /// 获取布尔字段的期待选择值。
    /// </summary>
    /// <param name="mode">筛选选择模式。</param>
    /// <returns>对应模式的布尔筛选项。</returns>
    private static object[] BooleanSelection(string mode) => mode switch
    {
        "single" => new object[] { false },
        "multiple" => new object[] { false, true },
        _ => Array.Empty<object>()
    };

    /// <summary>
    /// 获取空值字段的期待选择值。
    /// </summary>
    /// <param name="mode">筛选选择模式。</param>
    /// <returns>对应模式的空值筛选项。</returns>
    private static object[] OptionalSelection(string mode) => mode switch
    {
        "single" => new object[] { null },
        "multiple" => new object[] { null, "filled" },
        _ => Array.Empty<object>()
    };

    /// <summary>
    /// 获取状态字段的大小写不敏感筛选值。
    /// </summary>
    /// <param name="mode">筛选选择模式。</param>
    /// <returns>选择两个原始大小写项目的筛选值。</returns>
    private static object[] StatusSelection(string mode) => mode == "all"
        ? Array.Empty<object>() : new object[] { "rEtAiL" };

    /// <summary>
    /// 读取压缩包内的 XML 部件。
    /// </summary>
    /// <param name="archive">XLSX 包。</param>
    /// <param name="partName">部件路径。</param>
    /// <returns>解析后的 XML 文档。</returns>
    private static XDocument ReadXml(ZipArchive archive, string partName)
    {
        using var stream = archive.GetEntry(partName).Open();
        return XDocument.Load(stream);
    }
}
