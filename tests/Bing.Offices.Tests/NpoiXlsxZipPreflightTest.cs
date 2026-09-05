using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Imports;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// XLSX ZIP 预检安全边界测试。
/// </summary>
public sealed class NpoiXlsxZipPreflightTest
{
    /// <summary>
    /// 测试 - ZIP 压缩比限制显式关闭时，资源限制配置应通过校验。
    /// </summary>
    [Fact]
    public void ResourceLimits_NullCompressionRatio_ShouldBeAllowed()
    {
        // Arrange
        var limits = new ExcelResourceLimits { MaxZipCompressionRatio = null };

        // Act
        limits.Validate();

        // Assert
        Assert.Null(limits.MaxZipCompressionRatio);
    }

    /// <summary>
    /// 测试 - ZIP 压缩比限制为非正数或非有限值时应拒绝配置。
    /// </summary>
    [Fact]
    public void ResourceLimits_InvalidCompressionRatio_ShouldBeRejected()
    {
        // Arrange
        var invalidValues = new[] { 0d, -1d, double.NaN, double.PositiveInfinity, double.NegativeInfinity };

        // Act / Assert
        foreach (var value in invalidValues)
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                new ExcelResourceLimits { MaxZipCompressionRatio = value }.Validate());
    }

    /// <summary>
    /// 测试 - XML 字符和深度预算必须拒绝非正配置。
    /// </summary>
    [Fact]
    public void ResourceLimits_InvalidXmlBudgets_ShouldBeRejected()
    {
        // Act / Assert
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelResourceLimits { MaxXmlCharacters = 0 }.Validate());
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelResourceLimits { MaxXmlDepth = 0 }.Validate());
    }

    /// <summary>
    /// 测试 - XML 字符和深度限制显式关闭时，正常 XML 应继续通过预检。
    /// </summary>
    [Fact]
    public void ResourceLimits_NullXmlBudgets_ShouldBeAllowed()
    {
        // Arrange
        var limits = new ExcelResourceLimits
        {
            MaxXmlCharacters = null,
            MaxXmlDepth = null,
            MaxZipCompressionRatio = null
        };

        // Act
        limits.Validate();
        using var source = CreateZip(("xl/workbook.xml", "<workbook><a><b /></a></workbook>"));
        NpoiXlsxZipPreflight.Validate(source, limits);

        // Assert
        Assert.Null(limits.MaxXmlCharacters);
        Assert.Null(limits.MaxXmlDepth);
    }

    /// <summary>
    /// 测试 - 正常包含 workbook.xml 的 ZIP 应通过预检并保持源流位置可复用。
    /// </summary>
    [Fact]
    public void Validate_NormalXlsxZip_ShouldPassAndResetPosition()
    {
        // Arrange
        using var source = CreateZip(("xl/workbook.xml", "<workbook />"));
        source.Position = source.Length;

        // Act
        NpoiXlsxZipPreflight.Validate(source, new ExcelResourceLimits());

        // Assert
        Assert.Equal(0, source.Position);
    }

    /// <summary>
    /// 测试 - entry 数量超过预算时应在 NPOI DOM 创建前抛出资源限制异常。
    /// </summary>
    [Fact]
    public void Validate_EntryCountExceeded_ShouldThrowResourceLimitException()
    {
        // Arrange
        using var source = CreateZip(("xl/workbook.xml", "<workbook />"), ("xl/worksheets/sheet1.xml", "<worksheet />"));
        var limits = new ExcelResourceLimits { MaxZipEntries = 1 };

        // Act
        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiXlsxZipPreflight.Validate(source, limits));

        // Assert
        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal("NPOI", exception.Provider);
    }

    /// <summary>
    /// 测试 - 单个 XML entry 解压后超过预算时应拒绝输入。
    /// </summary>
    [Fact]
    public void Validate_EntrySizeExceeded_ShouldThrowResourceLimitException()
    {
        // Arrange
        using var source = CreateZip(("xl/workbook.xml", new string('x', 256)));
        var limits = new ExcelResourceLimits { MaxZipEntryUncompressedBytes = 32 };

        // Act
        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiXlsxZipPreflight.Validate(source, limits));

        // Assert
        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
    }

    /// <summary>
    /// 测试 - 所有 ZIP entry 的解压总量超过预算时应拒绝输入。
    /// </summary>
    [Fact]
    public void Validate_TotalUncompressedSizeExceeded_ShouldThrowResourceLimitException()
    {
        // Arrange
        using var source = CreateZip(("xl/workbook.xml", "<workbook />"),
            ("xl/worksheets/sheet1.xml", new string('x', 256)));
        var limits = new ExcelResourceLimits { MaxZipTotalUncompressedBytes = 64 };

        // Act
        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiXlsxZipPreflight.Validate(source, limits));

        // Assert
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("总解压大小", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 测试 - 高压缩比 entry 超过预算时应拒绝输入。
    /// </summary>
    [Fact]
    public void Validate_CompressionRatioExceeded_ShouldThrowResourceLimitException()
    {
        // Arrange
        using var source = CreateZip(("xl/workbook.xml", new string('x', 4096)));
        var limits = new ExcelResourceLimits { MaxZipCompressionRatio = 1 };

        // Act
        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiXlsxZipPreflight.Validate(source, limits));

        // Assert
        Assert.Contains("压缩比", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 测试 - 各类 XLSX XML 部件的字符预算超过限制时，应在 NPOI DOM 创建前拒绝输入。
    /// </summary>
    [Theory]
    [InlineData("xl/workbook.xml")]
    [InlineData("xl/sharedStrings.xml")]
    [InlineData("xl/styles.xml")]
    [InlineData("xl/worksheets/sheet1.xml")]
    public void Validate_XmlCharacterBudgetExceeded_ShouldThrowResourceLimitException(string entryName)
    {
        // Arrange
        using var source = CreateXmlBudgetZip(entryName,
            entryName == "xl/workbook.xml"
                ? "<workbook>1234567890123456789012345678901234567890</workbook>"
                : "<root>1234567890123456789012345678901234567890</root>");
        var limits = new ExcelResourceLimits
        {
            MaxXmlCharacters = 32,
            MaxZipCompressionRatio = null
        };

        // Act
        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiXlsxZipPreflight.Validate(source, limits));

        // Assert
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("字符数量", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 测试 - 各类 XLSX XML 部件的嵌套深度超过限制时，应在 NPOI DOM 创建前拒绝输入。
    /// </summary>
    [Theory]
    [InlineData("xl/workbook.xml")]
    [InlineData("xl/sharedStrings.xml")]
    [InlineData("xl/styles.xml")]
    [InlineData("xl/worksheets/sheet1.xml")]
    public void Validate_XmlDepthExceeded_ShouldThrowResourceLimitException(string entryName)
    {
        // Arrange
        using var source = CreateXmlBudgetZip(entryName,
            entryName == "xl/workbook.xml"
                ? "<workbook><a><b><c /></b></a></workbook>"
                : "<root><a><b><c /></b></a></root>");
        var limits = new ExcelResourceLimits
        {
            MaxXmlDepth = 2,
            MaxZipCompressionRatio = null
        };

        // Act
        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiXlsxZipPreflight.Validate(source, limits));

        // Assert
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("嵌套深度", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 测试 - XML 深度恰好达到配置上限时应通过预检。
    /// </summary>
    [Fact]
    public void Validate_XmlDepthAtLimit_ShouldPass()
    {
        // Arrange
        using var source = CreateZip(("xl/workbook.xml", "<workbook><a><b /></a></workbook>"));
        var limits = new ExcelResourceLimits
        {
            MaxXmlDepth = 2,
            MaxZipCompressionRatio = null
        };

        // Act
        NpoiXlsxZipPreflight.Validate(source, limits);

        // Assert
        Assert.Equal(0, source.Position);
    }

    /// <summary>
    /// 测试 - sharedStrings、styles 和单个 worksheet 部件分别受专项预算约束。
    /// </summary>
    [Theory]
    [InlineData("xl/sharedStrings.xml", "sharedStrings.xml")]
    [InlineData("xl/styles.xml", "styles.xml")]
    [InlineData("xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml")]
    public void Validate_SpecializedXmlSizeExceeded_ShouldThrowResourceLimitException(string entryName,
        string expectedName)
    {
        // Arrange
        using var source = CreateZip(("xl/workbook.xml", "<workbook />"),
            (entryName, new string('x', 256)));
        var limits = new ExcelResourceLimits
        {
            MaxSharedStringsBytes = entryName == "xl/sharedStrings.xml" ? 32 : null,
            MaxStylesBytes = entryName == "xl/styles.xml" ? 32 : null,
            MaxWorksheetBytes = entryName.StartsWith("xl/worksheets/", StringComparison.Ordinal)
                ? 32 : null
        };

        // Act
        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiXlsxZipPreflight.Validate(source, limits));

        // Assert
        Assert.Contains(expectedName, exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 测试 - 多个 worksheet 的 XML 总量超过专项总预算时应拒绝输入。
    /// </summary>
    [Fact]
    public void Validate_TotalWorksheetSizeExceeded_ShouldThrowResourceLimitException()
    {
        // Arrange
        using var source = CreateZip(("xl/workbook.xml", "<workbook />"),
            ("xl/worksheets/sheet1.xml", $"<worksheet>{new string('x', 64)}</worksheet>"),
            ("xl/worksheets/sheet2.xml", $"<worksheet>{new string('y', 64)}</worksheet>"));
        var limits = new ExcelResourceLimits
        {
            MaxWorksheetBytes = 1024,
            MaxTotalWorksheetBytes = 96
        };

        // Act
        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiXlsxZipPreflight.Validate(source, limits));

        // Assert
        Assert.Contains("worksheet XML 总大小", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 测试 - 重复 entry 和异常路径均应在 DOM 创建前被拒绝。
    /// </summary>
    [Fact]
    public void Validate_DuplicateOrInvalidEntryPath_ShouldThrowResourceLimitException()
    {
        // Arrange
        using var duplicate = CreateZip(("xl/workbook.xml", "<workbook />"),
            ("xl/workbook.xml", "<workbook />"));
        using var invalidPath = CreateZip(("xl/workbook.xml", "<workbook />"),
            ("../outside.xml", "<outside />"));

        // Act
        var duplicateException = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiXlsxZipPreflight.Validate(duplicate, new ExcelResourceLimits()));
        var pathException = Assert.Throws<BingOfficesResourceLimitException>(() =>
            NpoiXlsxZipPreflight.Validate(invalidPath, new ExcelResourceLimits()));

        // Assert
        Assert.Contains("重复 entry", duplicateException.Message, StringComparison.Ordinal);
        Assert.Contains("路径无效", pathException.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 测试 - XML DTD 或实体声明不得通过 XLSX 预检。
    /// </summary>
    [Fact]
    public void Validate_DtdEntity_ShouldThrowImportException()
    {
        // Arrange
        using var source = CreateZip(("xl/workbook.xml",
            "<!DOCTYPE workbook [<!ENTITY value 'x'>]><workbook>&value;</workbook>"));

        // Act
        var exception = Assert.Throws<BingOfficesImportException>(() =>
            NpoiXlsxZipPreflight.Validate(source, new ExcelResourceLimits()));

        // Assert
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.NotNull(exception.InnerException);
    }

    /// <summary>
    /// 测试 - 已取消请求应在读取 ZIP 或构造 Workbook 前原样传播取消异常。
    /// </summary>
    [Fact]
    public void Validate_CancellationRequested_ShouldThrowOperationCanceledException()
    {
        // Arrange
        using var source = CreateZip(("xl/workbook.xml", "<workbook />"));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        // Act
        var exception = Assert.Throws<OperationCanceledException>(() =>
            NpoiXlsxZipPreflight.Validate(source, new ExcelResourceLimits(), cancellation.Token));

        // Assert
        Assert.Equal(cancellation.Token, exception.CancellationToken);
    }

    /// <summary>
    /// 测试 - 损坏 ZIP 或 XML 不应泄漏底层预检异常类型。
    /// </summary>
    [Fact]
    public void Validate_CorruptZipOrXml_ShouldThrowBingOfficesException()
    {
        // Arrange
        using var corruptZip = new MemoryStream(Encoding.UTF8.GetBytes("PK\u0003\u0004broken"));
        using var corruptXml = CreateZip(("xl/workbook.xml", "<workbook>"));

        // Act
        var zipException = Assert.Throws<BingOfficesImportException>(() =>
            NpoiXlsxZipPreflight.Validate(corruptZip, new ExcelResourceLimits()));
        var xmlException = Assert.Throws<BingOfficesImportException>(() =>
            NpoiXlsxZipPreflight.Validate(corruptXml, new ExcelResourceLimits()));

        // Assert
        Assert.Equal(BingOfficesStage.Preflight, zipException.Stage);
        Assert.Equal(BingOfficesStage.Preflight, xmlException.Stage);
        Assert.NotNull(zipException.InnerException);
        Assert.NotNull(xmlException.InnerException);
    }

    /// <summary>
    /// 测试 - 缺失 workbook.xml 的 ZIP 应被识别为无效 XLSX。
    /// </summary>
    [Fact]
    public void Validate_MissingWorkbookEntry_ShouldThrowImportException()
    {
        // Arrange
        using var source = CreateZip(("xl/worksheets/sheet1.xml", "<worksheet />"));

        // Act
        var exception = Assert.Throws<BingOfficesImportException>(() =>
            NpoiXlsxZipPreflight.Validate(source, new ExcelResourceLimits()));

        // Assert
        Assert.Equal(BingOfficesErrorCode.ImportFailed, exception.Code);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    private static MemoryStream CreateZip(params (string Name, string Content)[] entries)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            foreach (var entry in entries)
            {
                using var writer = new StreamWriter(archive.CreateEntry(entry.Name).Open(), Encoding.UTF8, 1024, false);
                writer.Write(entry.Content);
            }
        }
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateXmlBudgetZip(string entryName, string content) =>
        entryName == "xl/workbook.xml"
            ? CreateZip((entryName, content))
            : CreateZip(("xl/workbook.xml", "<workbook />"), (entryName, content));
}
