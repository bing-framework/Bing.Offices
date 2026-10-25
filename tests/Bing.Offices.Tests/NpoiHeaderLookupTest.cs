using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>验证 NPOI 表头映射顺序、读取范围及请求间状态隔离。</summary>
public sealed class NpoiHeaderLookupTest
{
    /// <summary>跨属性标题、别名和属性名冲突时，较早映射应在两种文件格式中获胜。</summary>
    [Theory]
    [InlineData("xlsx", false)]
    [InlineData("xls", false)]
    [InlineData("xlsx", true)]
    [InlineData("xls", true)]
    public async Task Import_ConflictingHeaderTokens_ShouldUseMappingOrder(string format, bool asynchronous)
    {
        var importer = new NpoiExcelImporter();
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                Column(nameof(TokenConflictRow.First), "FirstTitle", 0, "TokenAliasName", "ThirdTitle"),
                Column(nameof(TokenConflictRow.TokenAliasName), "SecondTitle", 1, "FirstTitle"),
                Column(nameof(TokenConflictRow.Third), "ThirdTitle", 2)
            }
        };
        var request = CreateRequest<TokenConflictRow>(sheet => sheet.Mapping(mapping)
            .RequireExpectedHeaders(false));

        foreach (var conflictingHeader in new[] { "TokenAliasName", "FirstTitle", "ThirdTitle" })
        {
            var result = await Import(importer, CreateWorkbook(format, (0, conflictingHeader, "selected")),
                asynchronous, request);

            AssertSuccess(result);
            var item = Assert.Single(result.Workbook.Rows);
            Assert.Equal("selected", item.First);
            Assert.Null(item.TokenAliasName);
            Assert.Null(item.Third);
            Assert.Equal(new[] { 1 }, Assert.Single(result.Sheets).SourceRows);
        }
    }

    /// <summary>固定属性和动态列使用相同令牌时，固定属性在两种文件格式中优先。</summary>
    [Theory]
    [InlineData("xlsx", false)]
    [InlineData("xls", false)]
    [InlineData("xlsx", true)]
    [InlineData("xls", true)]
    public async Task Import_FixedAndDynamicTokenConflict_ShouldPreferFixedProperty(string format,
        bool asynchronous)
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                Column(nameof(FixedAndDynamicRow.Name), "Shared", 0)
            }
        };
        var request = CreateRequest<FixedAndDynamicRow>(sheet => sheet.Mapping(mapping)
            .DynamicColumns(row => row.Values, new[]
            {
                new ExcelDynamicColumnDefinition { Key = "shared", Title = "Shared" }
            })
            .RequireExpectedHeaders(false));

        var result = await Import(new NpoiExcelImporter(), CreateWorkbook(format, (0, "Shared", "fixed")),
            asynchronous, request);

        AssertSuccess(result);
        var item = Assert.Single(result.Workbook.Rows);
        Assert.Equal("fixed", item.Name);
        Assert.Empty(item.Values);
        Assert.Equal(new[] { 1 }, Assert.Single(result.Sheets).SourceRows);
    }

    /// <summary>比较器和映射连续变化时互不泄漏，失败后同一请求可再次成功。</summary>
    [Theory]
    [InlineData("xlsx", false)]
    [InlineData("xls", false)]
    [InlineData("xlsx", true)]
    [InlineData("xls", true)]
    public async Task Import_SequentialMappingsAndComparers_ShouldRemainIsolated(string format,
        bool asynchronous)
    {
        var importer = new NpoiExcelImporter();
        var ignoreCaseRequest = CreateRequest<ComparerIsolationRow>(sheet => sheet
            .Mapping(new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    Column(nameof(ComparerIsolationRow.Name), "Mapped", 0),
                    Column(nameof(ComparerIsolationRow.Code), "Other", 1)
                }
            })
            .HeaderComparison(ExcelNameComparison.OrdinalIgnoreCase));
        var caseInsensitive = await Import(importer,
            CreateWorkbook(format, (0, "mapped", "name"), (1, "Other", "other")), asynchronous,
            ignoreCaseRequest);
        AssertSuccess(caseInsensitive);
        var firstItem = Assert.Single(caseInsensitive.Workbook.Rows);
        Assert.Equal("name", firstItem.Name);
        Assert.Equal("other", firstItem.Code);

        var ordinalRequest = CreateRequest<ComparerIsolationRow>(sheet => sheet
            .Mapping(new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    Column(nameof(ComparerIsolationRow.Name), "Other", 0),
                    Column(nameof(ComparerIsolationRow.Code), "Mapped", 1)
                }
            })
            .HeaderComparison(ExcelNameComparison.Ordinal));
        var failed = await Import(importer,
            CreateWorkbook(format, (0, "Other", "other"), (1, "mapped", "lowercase")), asynchronous,
            ordinalRequest);
        AssertInvalidHeader(failed, "导入的表格不存在列：Mapped");

        var recovered = await Import(importer,
            CreateWorkbook(format, (0, "Other", "other"), (1, "Mapped", "code")), asynchronous,
            ordinalRequest);
        AssertSuccess(recovered);
        var secondItem = Assert.Single(recovered.Workbook.Rows);
        Assert.Equal("other", secondItem.Name);
        Assert.Equal("code", secondItem.Code);
    }

    /// <summary>必需列缺失时返回结构化错误，读取范围外的匹配列则豁免缺失检查。</summary>
    [Theory]
    [InlineData("xlsx", false)]
    [InlineData("xls", false)]
    [InlineData("xlsx", true)]
    [InlineData("xls", true)]
    public async Task Import_MissingRequiredHeaderAndOutsideRangeMatch_ShouldBehavePrecisely(string format,
        bool asynchronous)
    {
        var importer = new NpoiExcelImporter();
        var requiredRequest = CreateRequest<RequiredHeadersRow>(null);
        var missing = await Import(importer, CreateWorkbook(format, (0, "Name", "first")), asynchronous,
            requiredRequest);
        AssertInvalidHeader(missing, "导入的表格不存在列：Code");

        var rangedRequest = CreateRequest<RequiredHeadersRow>(sheet => sheet.ReadColumns(0, 1));
        var outsideRangeMatch = await Import(importer,
            CreateWorkbook(format, (0, "Name", "first"), (2, "Code", "outside")), asynchronous,
            rangedRequest);
        AssertSuccess(outsideRangeMatch);
        var item = Assert.Single(outsideRangeMatch.Workbook.Rows);
        Assert.Equal("first", item.Name);
        Assert.Null(item.Code);
        Assert.Equal(new[] { 1 }, Assert.Single(outsideRangeMatch.Sheets).SourceRows);
    }

    private static ExcelColumnConfiguration Column(string property, string title, int index,
        params string[] aliases) => new()
    {
        PropertyName = property,
        Title = title,
        ColumnIndex = index,
        Aliases = new List<string>(aliases)
    };

    private static ExcelWorkbookImportRequest<Workbook<TItem>> CreateRequest<TItem>(
        System.Action<ExcelSheetImportBuilder<TItem>> configure) where TItem : class, new() =>
        ExcelImport.Workbook<Workbook<TItem>>(builder => builder.Sheet("Data", root => root.Rows, configure));

    private static async Task<ExcelWorkbookImportResult<Workbook<TItem>>> Import<TItem>(NpoiExcelImporter importer,
        byte[] bytes, bool asynchronous, ExcelWorkbookImportRequest<Workbook<TItem>> request)
        where TItem : class, new()
    {
        using var source = new MemoryStream(bytes, writable: false);
        var result = asynchronous
            ? await importer.ImportAsync(source, request)
            : importer.Import(source, request);
        Assert.True(source.CanRead);
        return result;
    }

    private static byte[] CreateWorkbook(string format, params (int Column, string Header, string Value)[] cells) =>
        string.Equals(format, "xls", System.StringComparison.Ordinal)
            ? CreateWorkbook<HSSFWorkbook>(cells)
            : CreateWorkbook<XSSFWorkbook>(cells);

    private static byte[] CreateWorkbook<TWorkbook>(IReadOnlyList<(int Column, string Header, string Value)> cells)
        where TWorkbook : IWorkbook, new()
    {
        using var workbook = new TWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var header = sheet.CreateRow(0);
        var data = sheet.CreateRow(1);
        foreach (var cell in cells)
        {
            header.CreateCell(cell.Column).SetCellValue(cell.Header);
            data.CreateCell(cell.Column).SetCellValue(cell.Value);
        }
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    private static void AssertSuccess<TItem>(ExcelWorkbookImportResult<Workbook<TItem>> result)
        where TItem : class, new()
    {
        Assert.True(result.IsSuccess);
        Assert.Empty(result.Errors);
        Assert.False(result.ErrorsTruncated);
        Assert.Null(result.MaxErrors);
        var sheet = Assert.Single(result.Sheets);
        Assert.Equal("Data", sheet.Name);
        Assert.Empty(sheet.Errors);
    }

    private static void AssertInvalidHeader<TItem>(ExcelWorkbookImportResult<Workbook<TItem>> result,
        string message) where TItem : class, new()
    {
        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        Assert.False(result.ErrorsTruncated);
        Assert.Null(result.MaxErrors);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.InvalidHeader, error.Code);
        Assert.Equal(message, error.Message);
        Assert.Equal("Data", error.SheetName);
        Assert.Equal(1, error.RowIndex);
        Assert.Equal(0, error.ColumnIndex);
        Assert.Null(error.PropertyName);
        Assert.Equal(message, Assert.Single(Assert.Single(result.Sheets).Errors).Message);
    }

    private sealed class Workbook<TItem> where TItem : class, new()
    {
        public List<TItem> Rows { get; } = new();
    }

    private sealed class TokenConflictRow
    {
        public string First { get; set; }
        public string TokenAliasName { get; set; }
        public string Third { get; set; }
    }

    private sealed class FixedAndDynamicRow
    {
        public string Name { get; set; }

        [DynamicColumn]
        public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>();
    }

    private sealed class ComparerIsolationRow
    {
        public string Name { get; set; }
        public string Code { get; set; }
    }

    private sealed class RequiredHeadersRow
    {
        public string Name { get; set; }
        public string Code { get; set; }
    }
}
