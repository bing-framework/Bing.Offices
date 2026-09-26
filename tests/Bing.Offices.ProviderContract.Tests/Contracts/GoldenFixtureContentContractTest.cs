using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证冻结 Golden XLSX 的公式、日期和模板部件仍符合清单声明。
/// </summary>
public sealed class GoldenFixtureContentContractTest
{
    /// <summary>
    /// 定位工作簿 XML 元素使用的 SpreadsheetML 命名空间。
    /// </summary>
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// 公式 Golden 输入应保留公式文本和声明的缓存结果。
    /// </summary>
    /// <param name="relativePath">相对于黄金样例目录的文件路径。</param>
    /// <param name="formula">预期的公式文本。</param>
    /// <param name="cachedValue">预期缓存值文本。</param>
    /// <param name="cellType">预期的 OOXML 单元格类型标记；无标记时为 null。</param>
    [Theory]
    [InlineData("formula/cached-number.xlsx", "1+1", "2", null)]
    [InlineData("formula/cached-bool.xlsx", "1=1", "1", "b")]
    [InlineData("formula/cached-string.xlsx", "CONCAT(\"A\",\"B\")", "AB", "str")]
    [InlineData("formula/no-cache.xlsx", "1+1", "", null)]
    public void FormulaFixture_ShouldPreserveFormulaAndCachedValue(
        string relativePath, string formula, string cachedValue, string cellType)
    {
        var cell = WorksheetCell(relativePath, "A2");

        Assert.Equal(formula, cell.Element(Main + "f")?.Value);
        Assert.Equal(cachedValue, cell.Element(Main + "v")?.Value ?? string.Empty);
        Assert.Equal(cellType, (string)cell.Attribute("t"));
    }

    /// <summary>
    /// 日期 Golden 输入应保留 1900/1904 系统标记和日期样式。
    /// </summary>
    /// <param name="relativePath">相对于黄金样例目录的文件路径。</param>
    /// <param name="expectedDate1904">预期的 1904 日期系统标记。</param>
    /// <param name="expectedFormat">预期的日期格式代码。</param>
    [Theory]
    [InlineData("date/date1900.xlsx", "0", "yyyy-mm-dd")]
    [InlineData("date/date1904.xlsx", "1", "yyyy-mm-dd")]
    public void DateFixture_ShouldPreserveDateSystemAndStyle(
        string relativePath, string expectedDate1904, string expectedFormat)
    {
        using var archive = Open(relativePath);
        var workbook = XDocument.Parse(ReadEntry(archive, "xl/workbook.xml"));
        var workbookPr = workbook.Root?.Element(Main + "workbookPr");
        Assert.Equal(expectedDate1904, (string)workbookPr?.Attribute("date1904"));

        var styles = ReadEntry(archive, "xl/styles.xml");
        Assert.Contains($"formatCode=\"{expectedFormat}\"", styles, StringComparison.Ordinal);
        Assert.Equal("1", (string)WorksheetCell(archive, "A2").Attribute("s"));
    }

    /// <summary>
    /// 模板 Golden 输入应保留基础样式、批注、条件格式和合并区域。
    /// </summary>
    [Fact]
    public void TemplateFixtures_ShouldPreserveDeclaredParts()
    {
        using (var archive = Open("template/styles.xlsx"))
        {
            Assert.Equal("3", (string)WorksheetCell(archive, "A1").Attribute("s"));
            Assert.Equal("1", (string)WorksheetCell(archive, "B2").Attribute("s"));
        }

        using (var archive = Open("template/comments.xlsx"))
        {
            Assert.Contains("GoldenFixtures", ReadEntry(archive, "xl/comments1.xml"), StringComparison.Ordinal);
            Assert.Contains("legacyDrawing", ReadEntry(archive, "xl/worksheets/sheet1.xml"), StringComparison.Ordinal);
        }

        using (var archive = Open("template/conditional-formatting.xlsx"))
        {
            var sheet = XDocument.Parse(ReadEntry(archive, "xl/worksheets/sheet1.xml"));
            var conditional = sheet.Root?.Element(Main + "conditionalFormatting");
            Assert.Equal("A2:A3", (string)conditional?.Attribute("sqref"));
            Assert.Equal("A2>0", conditional?.Element(Main + "cfRule")?.Element(Main + "formula")?.Value);
        }

        using (var archive = Open("template/merged-cells.xlsx"))
        {
            var sheet = XDocument.Parse(ReadEntry(archive, "xl/worksheets/sheet1.xml"));
            Assert.Equal("A1:B1", (string)sheet.Root?.Element(Main + "mergeCells")?.Element(Main + "mergeCell")?.Attribute("ref"));
        }
    }

    /// <summary>
    /// 读取黄金样例第一个工作表中的指定单元格。
    /// </summary>
    /// <param name="relativePath">相对于黄金样例目录的文件路径。</param>
    /// <param name="reference">单元格的 A1 引用。</param>
    /// <returns>匹配指定引用的单元格 XML 元素。</returns>
    private static XElement WorksheetCell(string relativePath, string reference)
    {
        using var archive = Open(relativePath);
        return WorksheetCell(archive, reference);
    }

    /// <summary>
    /// 读取黄金样例第一个工作表中的指定单元格。
    /// </summary>
    /// <param name="archive">已打开的黄金 XLSX 压缩包。</param>
    /// <param name="reference">单元格的 A1 引用。</param>
    /// <returns>匹配指定引用的单元格 XML 元素。</returns>
    private static XElement WorksheetCell(ZipArchive archive, string reference)
    {
        var document = XDocument.Parse(ReadEntry(archive, "xl/worksheets/sheet1.xml"));
        var cell = document.Root?.Descendants(Main + "c")
            .SingleOrDefault(item => (string)item.Attribute("r") == reference);
        Assert.NotNull(cell);
        return cell!;
    }

    /// <summary>
    /// 以只读方式打开黄金 XLSX 样例。
    /// </summary>
    /// <param name="relativePath">相对于黄金样例目录的文件路径。</param>
    /// <returns>已打开的只读 ZIP 压缩包。</returns>
    private static ZipArchive Open(string relativePath)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Resources", "Golden",
            relativePath.Replace('/', Path.DirectorySeparatorChar));
        Assert.True(File.Exists(path), $"Golden input was not copied: {path}");
        return ZipFile.OpenRead(path);
    }

    /// <summary>
    /// 读取 ZIP 条目的完整文本。
    /// </summary>
    /// <param name="archive">已打开的黄金 XLSX 压缩包。</param>
    /// <param name="name">待读取的 ZIP 条目路径。</param>
    /// <returns>指定 ZIP 条目的完整文本。</returns>
    private static string ReadEntry(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry!.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }
}
