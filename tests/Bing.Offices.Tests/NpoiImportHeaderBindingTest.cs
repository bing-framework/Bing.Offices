using System;
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

/// <summary>
/// NPOI 导入表头目标绑定回归测试。
/// </summary>
public sealed class NpoiImportHeaderBindingTest
{
    /// <summary>
    /// 测试 - 标题、别名和 CLR 属性名同时出现时应拒绝重复绑定固定属性。
    /// </summary>
    [Theory]
    [InlineData("xlsx", false, "名称", "旧名称")]
    [InlineData("xlsx", false, "名称", "Name")]
    [InlineData("xlsx", false, "旧名称", "Name")]
    [InlineData("xls", false, "名称", "旧名称")]
    [InlineData("xls", false, "名称", "Name")]
    [InlineData("xls", false, "旧名称", "Name")]
    [InlineData("xlsx", true, "名称", "旧名称")]
    [InlineData("xlsx", true, "名称", "Name")]
    [InlineData("xlsx", true, "旧名称", "Name")]
    [InlineData("xls", true, "名称", "旧名称")]
    [InlineData("xls", true, "名称", "Name")]
    [InlineData("xls", true, "旧名称", "Name")]
    public async Task Import_DuplicateFixedHeaderTarget_ShouldReturnInvalidHeader(string format,
        bool asynchronous, string firstHeader, string secondHeader)
    {
        var workbook = CreateWorkbook(format, new[] { firstHeader, secondHeader }, new[] { "first", "second" });
        var request = CreateRequest(sheet => sheet.Mapping(CreateFixedMapping()));

        var result = await Import(workbook, asynchronous, request);

        AssertInvalidHeader(result);
    }

    /// <summary>
    /// 测试 - 动态列标题和别名绑定到同一声明键时应拒绝导入。
    /// </summary>
    [Theory]
    [InlineData("xlsx", false)]
    [InlineData("xls", false)]
    [InlineData("xlsx", true)]
    [InlineData("xls", true)]
    public async Task Import_DuplicateDeclaredDynamicHeaderTarget_ShouldReturnInvalidHeader(string format,
        bool asynchronous)
    {
        var workbook = CreateWorkbook(format, new[] { "名称", "地区", "旧地区" },
            new[] { "first", "east", "west" });
        var definitions = new[]
        {
            new ExcelDynamicColumnDefinition { Key = "region", Title = "地区", Aliases = new[] { "旧地区" } }
        };
        var request = CreateRequest(sheet => sheet.Mapping(CreateFixedMapping())
            .DynamicColumns(row => row.Values, definitions));

        var result = await Import(workbook, asynchronous, request);

        AssertInvalidHeader(result);
    }

    /// <summary>
    /// 测试 - 单个固定属性别名和不同声明动态列应完整导入，即使动态键与属性名相同。
    /// </summary>
    [Theory]
    [InlineData("xlsx", false)]
    [InlineData("xls", false)]
    [InlineData("xlsx", true)]
    [InlineData("xls", true)]
    public async Task Import_SingleAliasAndDistinctDeclaredDynamicColumns_ShouldImportValues(string format,
        bool asynchronous)
    {
        var workbook = CreateWorkbook(format, new[] { "旧名称", "旧地区", "省份" },
            new[] { "fixed value", "east", "north" });
        var definitions = new[]
        {
            new ExcelDynamicColumnDefinition { Key = "Name", Title = "地区", Aliases = new[] { "旧地区" } },
            new ExcelDynamicColumnDefinition { Key = "province", Title = "省份" }
        };
        var request = CreateRequest(sheet => sheet.Mapping(CreateFixedMapping())
            .DynamicColumns(row => row.Values, definitions));

        var result = await Import(workbook, asynchronous, request);

        Assert.True(result.IsSuccess);
        Assert.Empty(result.Errors);
        var item = Assert.Single(result.Workbook.Rows);
        Assert.Equal("fixed value", item.Name);
        Assert.Equal(2, item.Values.Count);
        Assert.Equal("east", item.Values["Name"]);
        Assert.Equal("north", item.Values["province"]);
        Assert.Equal(new[] { 1 }, Assert.Single(result.Sheets).SourceRows);
    }

    /// <summary>
    /// 测试 - 未声明动态列仍应按不同实际表头分别生成字典键。
    /// </summary>
    [Theory]
    [InlineData("xlsx", false)]
    [InlineData("xls", false)]
    [InlineData("xlsx", true)]
    [InlineData("xls", true)]
    public async Task Import_DifferentUnspecifiedDynamicHeaders_ShouldRemainDistinct(string format,
        bool asynchronous)
    {
        var workbook = CreateWorkbook(format, new[] { "名称", "临时甲", "临时乙" },
            new[] { "fixed value", "first", "second" });
        var request = CreateRequest(sheet => sheet.Mapping(CreateFixedMapping())
            .DynamicColumns(row => row.Values, Array.Empty<ExcelDynamicColumnDefinition>()));

        var result = await Import(workbook, asynchronous, request);

        Assert.True(result.IsSuccess);
        Assert.Empty(result.Errors);
        var item = Assert.Single(result.Workbook.Rows);
        Assert.Equal("fixed value", item.Name);
        Assert.Equal(2, item.Values.Count);
        Assert.Equal("first", item.Values["临时甲"]);
        Assert.Equal("second", item.Values["临时乙"]);
    }

    /// <summary>
    /// 测试 - CLR 名称仅大小写不同的固定属性应按标题比较规则分别绑定。
    /// </summary>
    [Theory]
    [InlineData("xlsx", false, ExcelNameComparison.OrdinalIgnoreCase)]
    [InlineData("xls", false, ExcelNameComparison.OrdinalIgnoreCase)]
    [InlineData("xlsx", true, ExcelNameComparison.OrdinalIgnoreCase)]
    [InlineData("xls", true, ExcelNameComparison.OrdinalIgnoreCase)]
    [InlineData("xlsx", false, ExcelNameComparison.Ordinal)]
    [InlineData("xls", false, ExcelNameComparison.Ordinal)]
    [InlineData("xlsx", true, ExcelNameComparison.Ordinal)]
    [InlineData("xls", true, ExcelNameComparison.Ordinal)]
    public async Task Import_CaseDistinctFixedProperties_ShouldBindBothValues(string format,
        bool asynchronous, ExcelNameComparison comparison)
    {
        var headers = comparison == ExcelNameComparison.Ordinal
            ? new[] { "Upper", "Lower" }
            : new[] { "upper", "lower" };
        var workbook = CreateWorkbook(format, headers, new[] { "first", "second" });
        var request = ExcelImport.Workbook<CaseDistinctNameWorkbook>(builder =>
            builder.Sheet("Data", root => root.Rows,
                sheet => sheet.HeaderComparison(comparison)));
        var importer = new NpoiExcelImporter();
        using var source = new MemoryStream(workbook, writable: false);

        var result = asynchronous
            ? await importer.ImportAsync(source, request)
            : importer.Import(source, request);

        Assert.True(result.IsSuccess);
        Assert.Empty(result.Errors);
        var item = Assert.Single(result.Workbook.Rows);
        Assert.Equal("first", item.Name);
        Assert.Equal("second", item.name);
        Assert.Equal(new[] { 1 }, Assert.Single(result.Sheets).SourceRows);
        Assert.True(source.CanRead);
        source.Position = 0;
        Assert.Equal(workbook[0], source.ReadByte());
    }

    private static void AssertInvalidHeader(ExcelWorkbookImportResult<HeaderBindingWorkbook> result)
    {
        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.InvalidHeader, error.Code);
        Assert.Equal("Data", error.SheetName);
        Assert.Equal(1, error.RowIndex);
        Assert.Equal(0, error.ColumnIndex);
        Assert.Null(error.PropertyName);
        var sheet = Assert.Single(result.Sheets);
        Assert.Equal("Data", sheet.Name);
        Assert.Empty(sheet.SourceRows);
        Assert.Equal(ExcelImportErrorCode.InvalidHeader, Assert.Single(sheet.Errors).Code);
    }

    private static ExcelMappingConfiguration CreateFixedMapping() => new()
    {
        Columns = new List<ExcelColumnConfiguration>
        {
            new()
            {
                PropertyName = nameof(HeaderBindingRow.Name),
                Title = "名称",
                Aliases = new List<string> { "旧名称" }
            }
        }
    };

    private static ExcelWorkbookImportRequest<HeaderBindingWorkbook> CreateRequest(
        Action<ExcelSheetImportBuilder<HeaderBindingRow>> configure) =>
        ExcelImport.Workbook<HeaderBindingWorkbook>(builder =>
            builder.Sheet("Data", root => root.Rows, configure));

    private static async Task<ExcelWorkbookImportResult<HeaderBindingWorkbook>> Import(byte[] workbook,
        bool asynchronous, ExcelWorkbookImportRequest<HeaderBindingWorkbook> request)
    {
        var importer = new NpoiExcelImporter();
        using var source = new MemoryStream(workbook, writable: false);
        return asynchronous
            ? await importer.ImportAsync(source, request)
            : importer.Import(source, request);
    }

    private static byte[] CreateWorkbook(string format, IReadOnlyList<string> headers,
        IReadOnlyList<string> values) => string.Equals(format, "xls", StringComparison.Ordinal)
        ? CreateWorkbook<HSSFWorkbook>(workbook => AddData(workbook, headers, values))
        : CreateWorkbook<XSSFWorkbook>(workbook => AddData(workbook, headers, values));

    private static byte[] CreateWorkbook<TWorkbook>(Action<TWorkbook> configure)
        where TWorkbook : IWorkbook, new()
    {
        using var workbook = new TWorkbook();
        configure(workbook);
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    private static void AddData(IWorkbook workbook, IReadOnlyList<string> headers, IReadOnlyList<string> values)
    {
        var sheet = workbook.CreateSheet("Data");
        var headerRow = sheet.CreateRow(0);
        var dataRow = sheet.CreateRow(1);
        for (var column = 0; column < headers.Count; column++)
        {
            headerRow.CreateCell(column).SetCellValue(headers[column]);
            dataRow.CreateCell(column).SetCellValue(values[column]);
        }
    }

    private sealed class HeaderBindingWorkbook
    {
        public List<HeaderBindingRow> Rows { get; } = new();
    }

    private sealed class HeaderBindingRow
    {
        public string Name { get; set; }

        [DynamicColumn]
        public IDictionary<string, object> Values { get; set; }
    }

    private sealed class CaseDistinctNameWorkbook
    {
        public List<CaseDistinctNameRow> Rows { get; } = new();
    }

    private sealed class CaseDistinctNameRow
    {
        [ColumnName("Upper")]
        public string Name { get; set; }

        [ColumnName("Lower")]
        public string name { get; set; }
    }
}
