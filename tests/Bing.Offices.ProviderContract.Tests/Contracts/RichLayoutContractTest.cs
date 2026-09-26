using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.ProviderContract.Tests.Profiles;
using Bing.Offices.Providers;
using Bing.Offices.Styles;
using Bing.Offices.Testing.Models;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 通过公共 API 和 Provider 无关的 OOXML 观察验证富布局合同。
/// </summary>
public sealed class RichLayoutContractTest
{
    /// <summary>
    /// 用于定位工作簿、工作表和样式元素的 OOXML 主命名空间。
    /// </summary>
    private static readonly XNamespace Main =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    /// <summary>
    /// 用于解析工作表关系标识的文档关系命名空间。
    /// </summary>
    private static readonly XNamespace OfficeRelationships =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    /// <summary>
    /// 用于查找工作簿部件关联的包关系命名空间。
    /// </summary>
    private static readonly XNamespace PackageRelationships =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// 获取参与富布局合同的提供程序名称集合。
    /// </summary>
    public static IEnumerable<object[]> Providers =>
        ProviderDrivers.All.Select(driver => new object[] { driver.Name });

    /// <summary>
    /// 验证样式输出符合独立能力预期和 OOXML 格式。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public void StyleContract_ShouldMatchProfileAndOoxml(string provider)
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new RichContractRow { Name = "A", Amount = 12.5m } }, sheet => sheet
                .HeaderStyle(new ExcelCellStyle { Bold = true })
                .BodyStyle(new ExcelCellStyle { NumberFormat = "0.00" })));
        using var output = new MemoryStream();

        if (!ExportOrAssertUnsupported(provider, ContractScenario.Style, request, output))
            return;

        using var archive = Open(output);
        var worksheet = Worksheet(archive, "Data");
        var styles = Xml(archive, "xl/styles.xml");
        AssertBold(styles, Cell(worksheet, "A1"));
        AssertNumberFormat(styles, Cell(worksheet, "B2"), "0.00", 2);
    }

    /// <summary>
    /// 验证合并表头符合独立能力预期和 OOXML 格式。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public void MergeContract_ShouldMatchProfileAndOoxml(string provider)
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new RichContractRow { Name = "A", Amount = 12.5m } }, sheet => sheet
                .HeaderRowIndex(1)
                .DataRowStartIndex(2)
                .HeaderRows(new[]
                {
                    new ExcelHeaderRow(0, new[] { new ExcelHeaderCell(0, "Summary", columnSpan: 2) })
                })));
        using var output = new MemoryStream();

        if (!ExportOrAssertUnsupported(provider, ContractScenario.Merge, request, output))
            return;

        using var archive = Open(output);
        var worksheet = Worksheet(archive, "Data");
        var merge = Assert.Single(worksheet.Descendants(Main + "mergeCell"));
        Assert.Equal("A1:B1", (string)merge.Attribute("ref"));
    }

    /// <summary>
    /// 验证模板导出保留粗体和数字格式。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public void TemplateContract_ShouldPreserveDeclaredStyles(string provider)
    {
        using var template = Golden("template/styles.xlsx");
        var original = template.ToArray();
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", new[] { new RichContractRow { Name = "Updated", Amount = 42.5m } }));
        using var output = new MemoryStream();

        if (!ExportOrAssertUnsupported(provider, ContractScenario.Template, request, output))
        {
            Assert.True(template.CanRead);
            return;
        }

        Assert.True(template.CanRead);
        Assert.Equal(original, template.ToArray());
        using var archive = Open(output);
        var worksheet = Worksheet(archive, "Data");
        var styles = Xml(archive, "xl/styles.xml");
        AssertBold(styles, Cell(worksheet, "A1"));
        AssertNumberFormat(styles, Cell(worksheet, "B2"), "yyyy-mm-dd");
    }

    /// <summary>
    /// 验证实体布局按声明能力完成公共接口往返。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public void EntityContract_ShouldMatchProfileAndRoundTrip(string provider)
    {
        var driver = ProviderDrivers.Get(provider);
        var expectation = ProviderContractProfiles.Get(provider).For(ContractScenario.Entity);
        if (expectation == ContractExpectation.NotApplicable)
        {
            Assert.Equal("MiniExcel", provider);
            Assert.False(driver.DeclaredCapabilities.Supports(ExcelProviderCapabilities.Entity));
            return;
        }

        Assert.Equal(ContractExpectation.Supported, expectation);
        var layout = ExcelEntity.Layout<EntityContractRoot>(builder => builder
            .Merge("Invoice", "A1:B1")
            .Cell("Invoice", "A1", entity => entity.Title)
            .Cell("Invoice", "B2", entity => entity.Number)
            .Cell("Invoice", "B3", entity => entity.Customer)
            .ListRegion("Lines", "A1", entity => entity.Lines));
        var source = new EntityContractRoot
        {
            Title = "Order",
            Number = 1001,
            Customer = "Customer A",
            Lines = new List<EntityContractLine>
            {
                new() { Code = "A", Quantity = 2 },
                new() { Code = "B", Quantity = 3 }
            }
        };
        using var output = new MemoryStream();

        var exporter = Assert.IsAssignableFrom<IExcelEntityExporter>(driver.CreateExporter());
        exporter.ExportEntity(source, layout, output);

        using (var archive = Open(output))
        {
            var worksheet = Worksheet(archive, "Invoice");
            var merge = Assert.Single(worksheet.Descendants(Main + "mergeCell"));
            Assert.Equal("A1:B1", (string)merge.Attribute("ref"));
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var importer = Assert.IsAssignableFrom<IExcelEntityImporter>(driver.CreateImporter());
        var result = importer.ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Order", result.Entity.Title);
        Assert.Equal(1001, result.Entity.Number);
        Assert.Equal("Customer A", result.Entity.Customer);
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(new[] { 2, 3 }, result.Entity.Lines.Select(line => line.Quantity).ToArray());
    }

    /// <summary>
    /// 执行导出或断言预期的不支持错误。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    /// <param name="scenario">独立能力预期对应的合同场景。</param>
    /// <param name="request">待执行的工作簿导出请求。</param>
    /// <param name="output">接收导出内容的目标流。</param>
    /// <returns>导出成功时为 true；确认预期的不支持错误时为 false。</returns>
    private static bool ExportOrAssertUnsupported(string provider, ContractScenario scenario,
        ExcelWorkbookExportRequest request, MemoryStream output)
    {
        var expectation = ProviderContractProfiles.Get(provider).For(scenario);
        var exporter = ProviderDrivers.Get(provider).CreateExporter();
        if (expectation == ContractExpectation.UnsupportedExpected)
        {
            var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
                exporter.Export(request, output));
            Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
            Assert.Equal(BingOfficesOperation.Export, exception.Operation);
            Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
            Assert.Equal(provider, exception.Provider);
            Assert.Equal(0, output.Length);
            return false;
        }

        Assert.Equal(ContractExpectation.Supported, expectation);
        exporter.Export(request, output);
        Assert.True(output.Length > 0);
        return true;
    }

    /// <summary>
    /// 读取独立黄金样本工作簿。
    /// </summary>
    /// <param name="relativePath">相对于黄金样本目录的文件路径。</param>
    /// <returns>包含黄金样本字节的只读内存流。</returns>
    private static MemoryStream Golden(string relativePath)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Resources", "Golden",
            relativePath.Replace('/', Path.DirectorySeparatorChar));
        Assert.True(File.Exists(path), $"Golden input was not copied: {path}");
        return new MemoryStream(File.ReadAllBytes(path), writable: false);
    }

    /// <summary>
    /// 从输出流副本打开只读 OOXML 包。
    /// </summary>
    /// <param name="stream">包含导出工作簿的内存流。</param>
    /// <returns>持有独立输出副本的只读 ZIP 包。</returns>
    private static ZipArchive Open(MemoryStream stream) =>
        new ZipArchive(new MemoryStream(stream.ToArray(), writable: false), ZipArchiveMode.Read);

    /// <summary>
    /// 按名称定位工作表 XML 部件。
    /// </summary>
    /// <param name="archive">待读取的 OOXML 包。</param>
    /// <param name="name">工作表或属性名称。</param>
    /// <returns>指定工作表的 XML 文档。</returns>
    private static XDocument Worksheet(ZipArchive archive, string name)
    {
        var workbook = Xml(archive, "xl/workbook.xml");
        var sheet = Assert.Single(workbook.Descendants(Main + "sheet")
            .Where(item => string.Equals((string)item.Attribute("name"), name, StringComparison.Ordinal)));
        var relationshipId = (string)sheet.Attribute(OfficeRelationships + "id");
        var relationships = Xml(archive, "xl/_rels/workbook.xml.rels");
        var relationship = Assert.Single(relationships.Descendants(PackageRelationships + "Relationship")
            .Where(item => string.Equals((string)item.Attribute("Id"), relationshipId,
                StringComparison.Ordinal)));
        var target = ((string)relationship.Attribute("Target")).Replace('\\', '/');
        var path = target.StartsWith("/", StringComparison.Ordinal)
            ? target.Substring(1)
            : "xl/" + target.TrimStart('/');
        return Xml(archive, path);
    }

    /// <summary>
    /// 读取并解析指定 XML 包部件。
    /// </summary>
    /// <param name="archive">待读取的 OOXML 包。</param>
    /// <param name="path">包内 XML 部件路径。</param>
    /// <returns>解析后的 XML 文档。</returns>
    private static XDocument Xml(ZipArchive archive, string path) =>
        XDocument.Parse(ReadEntry(archive, path));

    /// <summary>
    /// 读取指定包部件的 UTF-8 文本。
    /// </summary>
    /// <param name="archive">待读取的 OOXML 包。</param>
    /// <param name="path">包内 XML 部件路径。</param>
    /// <returns>包部件的完整文本。</returns>
    private static string ReadEntry(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path);
        Assert.NotNull(entry);
        using var reader = new StreamReader(entry!.Open(), Encoding.UTF8, true);
        return reader.ReadToEnd();
    }

    /// <summary>
    /// 按 A1 引用定位唯一单元格元素。
    /// </summary>
    /// <param name="worksheet">工作表 XML 文档。</param>
    /// <param name="reference">目标单元格的 A1 引用。</param>
    /// <returns>唯一匹配 A1 引用的单元格元素。</returns>
    private static XElement Cell(XDocument worksheet, string reference) =>
        Assert.Single(worksheet.Descendants(Main + "c")
            .Where(item => string.Equals((string)item.Attribute("r"), reference,
                StringComparison.Ordinal)));

    /// <summary>
    /// 断言单元格应用粗体字体。
    /// </summary>
    /// <param name="styles">工作簿样式 XML 文档。</param>
    /// <param name="cell">待验证的单元格元素。</param>
    private static void AssertBold(XDocument styles, XElement cell)
    {
        var cellFormat = CellFormat(styles, cell);
        var fontIndex = AttributeInt(cellFormat, "fontId");
        var fonts = styles.Descendants(Main + "fonts").Elements(Main + "font").ToArray();
        Assert.InRange(fontIndex, 0, fonts.Length - 1);
        Assert.NotNull(fonts[fontIndex].Element(Main + "b"));
    }

    /// <summary>
    /// 断言单元格应用预期数字格式。
    /// </summary>
    /// <param name="styles">工作簿样式 XML 文档。</param>
    /// <param name="cell">待验证的单元格元素。</param>
    /// <param name="expected">预期的自定义数字格式文本。</param>
    /// <param name="acceptedBuiltInId">可直接接受的内置数字格式编号；null 表示只验证自定义格式。</param>
    private static void AssertNumberFormat(XDocument styles, XElement cell, string expected,
        int? acceptedBuiltInId = null)
    {
        var cellFormat = CellFormat(styles, cell);
        var numberFormatId = AttributeInt(cellFormat, "numFmtId");
        if (acceptedBuiltInId == numberFormatId)
            return;
        var numberFormat = Assert.Single(styles.Descendants(Main + "numFmt")
            .Where(item => AttributeInt(item, "numFmtId") == numberFormatId));
        Assert.Equal(expected, (string)numberFormat.Attribute("formatCode"));
    }

    /// <summary>
    /// 读取单元格引用的样式记录。
    /// </summary>
    /// <param name="styles">工作簿样式 XML 文档。</param>
    /// <param name="cell">待验证的单元格元素。</param>
    /// <returns>单元格样式索引对应的格式记录。</returns>
    private static XElement CellFormat(XDocument styles, XElement cell)
    {
        var styleIndex = AttributeInt(cell, "s");
        var formats = styles.Descendants(Main + "cellXfs").Elements(Main + "xf").ToArray();
        Assert.InRange(styleIndex, 0, formats.Length - 1);
        return formats[styleIndex];
    }

    /// <summary>
    /// 读取必需的非负整数 XML 属性。
    /// </summary>
    /// <param name="element">包含目标属性的 XML 元素。</param>
    /// <param name="name">工作表或属性名称。</param>
    /// <returns>以固定区域性解析的整数属性值。</returns>
    private static int AttributeInt(XElement element, string name)
    {
        var value = (string)element.Attribute(name);
        Assert.False(string.IsNullOrWhiteSpace(value), $"Missing {name} on {element.Name.LocalName}.");
        return int.Parse(value!, NumberStyles.None, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 用于样式和模板合同验证的数据行。
    /// </summary>
    private sealed class RichContractRow
    {
        /// <summary>
        /// 获取或设置测试数据名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置用于数字格式验证的金额。
        /// </summary>
        public decimal Amount { get; set; }
    }
}
