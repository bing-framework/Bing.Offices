using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Extensions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI 与 MiniExcel 共同 XLSX 能力的结构化合同测试。
/// </summary>
public sealed class MiniExcelProviderContractTest
{
    /// <summary>
    /// 验证 NPOI 与 MiniExcel 在标量和动态列场景下的结果一致性。
    /// </summary>
    [Fact]
    public async Task CommonScalarAndDynamicContract_ShouldMatchAcrossProviders()
    {
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "Region",
            DataType = typeof(string)
        };
        var mapping = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(ContractRow.Code),
                    Title = "Code"
                },
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(ContractRow.Amount),
                    Title = "Amount",
                    Formatter = "0.00"
                }
            }
        };
        var rows = new[]
        {
            new ContractRow
            {
                Code = "A",
                Count = 7,
                Amount = 12.5m,
                Enabled = true,
                Kind = ContractKind.Active,
                Date = new DateTime(2026, 9, 4),
                Optional = null,
                Values = new Dictionary<string, object> { ["region"] = "east" }
            },
            new ContractRow
            {
                Code = "B",
                Count = 8,
                Amount = 3.25m,
                Enabled = false,
                Kind = ContractKind.Pending,
                Date = new DateTime(2026, 9, 5),
                Optional = 11,
                Values = new Dictionary<string, object> { ["region"] = "west" }
            }
        };
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows,
            sheet => sheet.Mapping(mapping)
                .DynamicColumns(row => row.Values, new[] { definition })));
        var importRequest = ExcelImport.Workbook<ContractWorkbook>(workbook => workbook
            .Sheet<ContractRow>("Data", root => root.Rows, sheet => sheet
                .Mapping(mapping)
                .DynamicColumns(row => row.Values, new[] { definition })));

        var npoi = await RoundTripAsync(new Bing.Offices.Exports.NpoiExcelExporter(),
            new Bing.Offices.Imports.NpoiExcelImporter(), exportRequest, importRequest);
        var mini = await RoundTripAsync(new MiniExcelExcelExporter(),
            new Bing.Offices.Imports.MiniExcelExcelImporter(), exportRequest, importRequest);

        Assert.True(npoi.IsSuccess, string.Join(";", npoi.Errors.Select(error => error.Message)));
        Assert.True(mini.IsSuccess, string.Join(";", mini.Errors.Select(error => error.Message)));
        Assert.Equal(npoi.Errors.Select(error => error.Code), mini.Errors.Select(error => error.Code));
        Assert.Equal(ToContractRows(npoi.Workbook.Rows), ToContractRows(mini.Workbook.Rows));
    }

    /// <summary>
    /// 验证 NPOI 与 MiniExcel 在值映射、转换器和校验场景下的结果一致性。
    /// </summary>
    [Fact]
    public async Task ValueMapConverterAndValidationContract_ShouldMatchAcrossProviders()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(RuleContractRow.Code),
                    Title = "Code"
                },
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(RuleContractRow.Status),
                    Title = "Status"
                },
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(RuleContractRow.Name),
                    Title = "Name",
                    ConverterName = "contract-upper"
                }
            }
        };
        var rows = new[]
        {
            new RuleContractRow { Code = "A", Status = 1, Name = "alice" },
            new RuleContractRow { Code = "B", Status = 1, Name = "bob" }
        };
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows,
            sheet => sheet.Mapping(mapping)));
        var importRequest = ExcelImport.Workbook<RuleContractWorkbook>(workbook => workbook
            .Sheet<RuleContractRow>("Data", root => root.Rows, sheet => sheet
                .Mapping(mapping).Validate(ExcelValidationFailureMode.Continue)));

        using var npoiStream = new System.IO.MemoryStream();
        await new Bing.Offices.Exports.NpoiExcelExporter(
                new IExcelValueConverter[] { new ContractUpperConverter() })
            .ExportAsync(exportRequest, npoiStream);
        npoiStream.Position = 0;
        var npoiConverter = new ContractUpperConverter();
        var npoi = await new Bing.Offices.Imports.NpoiExcelImporter(
                valueConverters: new IExcelValueConverter[] { npoiConverter })
            .ImportAsync(npoiStream, importRequest);

        using var miniStream = new System.IO.MemoryStream();
        await new MiniExcelExcelExporter(
                new IExcelValueConverter[] { new ContractUpperConverter() })
            .ExportAsync(exportRequest, miniStream);
        miniStream.Position = 0;
        var miniConverter = new ContractUpperConverter();
        var mini = await new Bing.Offices.Imports.MiniExcelExcelImporter(
                valueConverters: new IExcelValueConverter[] { miniConverter })
            .ImportAsync(miniStream, importRequest);

        Assert.True(npoi.IsSuccess, string.Join(";", npoi.Errors.Select(error => error.Message)));
        Assert.True(mini.IsSuccess, string.Join(";", mini.Errors.Select(error => error.Message)));
        Assert.Equal(ToRuleRows(npoi.Workbook.Rows), ToRuleRows(mini.Workbook.Rows));
        Assert.Equal(new[] { 1, 3 }, npoiConverter.ImportContexts.Select(context => context.ColumnIndex)
            .Distinct().ToArray());
        Assert.Equal(new[] { 1, 3 }, miniConverter.ImportContexts.Select(context => context.ColumnIndex)
            .Distinct().ToArray());
    }

    /// <summary>
    /// 验证 NPOI 与 MiniExcel 在结构化错误场景下返回一致的错误结果。
    /// </summary>
    [Fact]
    public async Task StructuredErrorContract_ShouldMatchAcrossProviders()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new RequiredContractRow { Code = string.Empty, Quantity = 0 } }));
        var importRequest = ExcelImport.Workbook<RequiredContractWorkbook>(workbook => workbook
            .Sheet<RequiredContractRow>("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

        var npoi = await RoundTripAsync(new Bing.Offices.Exports.NpoiExcelExporter(),
            new Bing.Offices.Imports.NpoiExcelImporter(), exportRequest, importRequest);
        var mini = await RoundTripAsync(new MiniExcelExcelExporter(),
            new Bing.Offices.Imports.MiniExcelExcelImporter(), exportRequest, importRequest);

        Assert.False(npoi.IsSuccess);
        Assert.False(mini.IsSuccess);
        Assert.Equal(ToErrorContract(npoi.Errors), ToErrorContract(mini.Errors));
        Assert.Empty(npoi.Workbook.Rows);
        Assert.Empty(mini.Workbook.Rows);
    }

    /// <summary>
    /// 验证 NPOI 与 MiniExcel 在关系映射场景下生成一致的关联结果。
    /// </summary>
    [Fact]
    public async Task RelationContract_ShouldMatchAcrossProviders()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Parents", new[] { new RelationContractParent { OrderNo = "A-1" } })
            .AddSheet("Children", new[]
            {
                new RelationContractChild { OrderNo = "a-1", Name = "Item-1" },
                new RelationContractChild { OrderNo = "A-1", Name = "Item-2" }
            }));
        var importRequest = ExcelImport.Workbook<RelationContractWorkbook>(workbook =>
        {
            workbook.Sheet("Parents", root => root.Parents);
            workbook.Sheet("Children", root => root.Children);
            workbook.HasMany(root => root.Parents, root => root.Children,
                parent => parent.OrderNo, child => child.OrderNo,
                parent => parent.Items, StringComparer.OrdinalIgnoreCase);
        });

        var npoi = await RoundTripAsync(new Bing.Offices.Exports.NpoiExcelExporter(),
            new Bing.Offices.Imports.NpoiExcelImporter(), exportRequest, importRequest);
        var mini = await RoundTripAsync(new MiniExcelExcelExporter(),
            new Bing.Offices.Imports.MiniExcelExcelImporter(), exportRequest, importRequest);

        Assert.True(npoi.IsSuccess, string.Join(";", npoi.Errors.Select(error => error.Message)));
        Assert.True(mini.IsSuccess, string.Join(";", mini.Errors.Select(error => error.Message)));
        Assert.Equal(ToRelationRows(npoi.Workbook), ToRelationRows(mini.Workbook));
        Assert.Equal(new[] { "Item-1", "Item-2" },
            npoi.Workbook.Parents.Single().Items.Select(item => item.Name).ToArray());
    }

    /// <summary>
    /// 验证不支持的功能会被两个 Provider 按约定处理。
    /// </summary>
    [Fact]
    public void UnsupportedFeatureContract_ShouldFailFastWithoutOutput()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Data", new[] { new ContractRow() }, sheet =>
                sheet.SheetStyle(new Bing.Offices.Styles.ExcelCellStyle { Bold = true })));
        using var npoiOutput = new System.IO.MemoryStream();
        using var miniOutput = new System.IO.MemoryStream();

        // NPOI 支持样式，MiniExcel 必须明确拒绝而不是静默生成不完整文件。
        new Bing.Offices.Exports.NpoiExcelExporter().Export(request, npoiOutput);
        var exception = Assert.Throws<Bing.Offices.Exceptions.BingOfficesUnsupportedFeatureException>(() =>
            new MiniExcelExcelExporter().Export(request, miniOutput));
        Assert.Equal(Bing.Offices.Exceptions.BingOfficesErrorCode.UnsupportedFeature, exception.Code);
        Assert.Equal(Bing.Offices.Exceptions.BingOfficesOperation.Export, exception.Operation);
        Assert.Equal("MiniExcel", exception.Provider);
        Assert.Equal(Bing.Offices.Exceptions.BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, miniOutput.Length);
        Assert.True(npoiOutput.Length > 0);
    }

    /// <summary>
    /// 验证 Provider 注册顺序保留首个注册的实现。
    /// </summary>
    [Fact]
    public void ProviderRegistrationOrder_ShouldPreserveFirstRegistration()
    {
        using var npoiFirst = new ServiceCollection()
            .AddBingOfficesNpoi()
            .AddBingOfficesMiniExcel()
            .BuildServiceProvider();
        using var miniExcelFirst = new ServiceCollection()
            .AddBingOfficesMiniExcel()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();

        Assert.IsType<Bing.Offices.Exports.NpoiExcelExporter>(
            npoiFirst.GetRequiredService<IExcelExporter>());
        Assert.IsType<Bing.Offices.Imports.NpoiExcelImporter>(
            npoiFirst.GetRequiredService<IExcelImporter>());
        Assert.IsType<MiniExcelExcelExporter>(
            miniExcelFirst.GetRequiredService<IExcelExporter>());
        Assert.IsType<Bing.Offices.Imports.MiniExcelExcelImporter>(
            miniExcelFirst.GetRequiredService<IExcelImporter>());
    }

    /// <summary>
    /// 验证首个注册的 Provider 可以完成公开往返流程。
    /// </summary>
    [Fact]
    public async Task ProviderRegistrationOrder_ShouldRunPublicRoundTripWithFirstProvider()
    {
        using var npoiFirst = new ServiceCollection()
            .AddBingOfficesNpoi()
            .AddBingOfficesMiniExcel()
            .BuildServiceProvider();
        using var miniExcelFirst = new ServiceCollection()
            .AddBingOfficesMiniExcel()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();

        await AssertDiRoundTripAsync(npoiFirst, "NPOI");
        await AssertDiRoundTripAsync(miniExcelFirst, "MiniExcel");
    }

    /// <summary>
    /// 断言依赖注入注册的异步往返结果。
    /// </summary>
    /// <param name="provider">Provider 实例。</param>
    /// <param name="expectedProvider">预期 Provider 类型。</param>
    private static async Task AssertDiRoundTripAsync(IServiceProvider provider, string expectedProvider)
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration { PropertyName = nameof(RegistrationRow.Code), Title = "Code" },
                new ExcelColumnConfiguration { PropertyName = nameof(RegistrationRow.Count), Title = "Count" }
            }
        };
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new RegistrationRow { Code = expectedProvider, Count = 7 } },
            sheet => sheet.Mapping(mapping)));
        var importRequest = ExcelImport.Workbook<RegistrationWorkbook>(workbook =>
            workbook.Sheet<RegistrationRow>("Data", root => root.Rows,
                sheet => sheet.Mapping(mapping)));
        using var stream = new System.IO.MemoryStream();
        await provider.GetRequiredService<IExcelExporter>().ExportAsync(exportRequest, stream);
        stream.Position = 0;
        var result = await provider.GetRequiredService<IExcelImporter>().ImportAsync(stream, importRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var row = Assert.Single(result.Workbook.Rows);
        Assert.Equal(expectedProvider, row.Code);
        Assert.Equal(7, row.Count);
    }

    /// <summary>
    /// 执行异步 Excel 往返并返回导入结果。
    /// </summary>
    /// <typeparam name="TWorkbook">泛型参数 TWorkbook 表示方法处理的数据类型。</typeparam>
    /// <param name="exporter">Excel 或 CSV 导出器。</param>
    /// <param name="importer">Excel 或 CSV 导入器。</param>
    /// <param name="exportRequest">导入或导出请求。</param>
    /// <param name="importRequest">导入或导出请求。</param>
    /// <returns>异步操作的最终结果。</returns>
    private static async Task<ExcelWorkbookImportResult<TWorkbook>> RoundTripAsync<TWorkbook>(
        IExcelExporter exporter, IExcelImporter importer,
        ExcelWorkbookExportRequest exportRequest,
        ExcelWorkbookImportRequest<TWorkbook> importRequest)
        where TWorkbook : class, new()
    {
        using var stream = new System.IO.MemoryStream();
        await exporter.ExportAsync(exportRequest, stream);
        stream.Position = 0;
        return await importer.ImportAsync(stream, importRequest);
    }

    /// <summary>
    /// 转换为合同测试行。
    /// </summary>
    /// <param name="rows">数据行或行集合。</param>
    /// <returns>按输入顺序格式化的合同数据行列表。</returns>
    private static IReadOnlyList<string> ToContractRows(IReadOnlyList<ContractRow> rows) => rows
        .Select(row => string.Join("|", row.Code, row.Count.ToString(CultureInfo.InvariantCulture),
            row.Amount.ToString("0.00", CultureInfo.InvariantCulture), row.Enabled,
            row.Kind, row.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            row.Optional?.ToString(CultureInfo.InvariantCulture) ?? "<null>",
            row.Values.TryGetValue("region", out var region) ? region?.ToString() : "<missing>"))
        .ToArray();

    /// <summary>
    /// 转换为规则测试行。
    /// </summary>
    /// <param name="rows">数据行或行集合。</param>
    /// <returns>包含转换和校验相关字段的可比较字符串列表。</returns>
    private static IReadOnlyList<string> ToRuleRows(IReadOnlyList<RuleContractRow> rows) => rows
        .Select(row => string.Join("|", row.Code, row.Status,
            row.Name, row.Name == row.Name.ToUpperInvariant()))
        .ToArray();

    /// <summary>
    /// 转换为错误合同数据。
    /// </summary>
    /// <param name="errors">错误集合。</param>
    /// <returns>包含错误代码、位置、属性和消息的可比较字符串列表。</returns>
    private static IReadOnlyList<string> ToErrorContract(IReadOnlyList<ExcelImportError> errors) => errors
        .Select(error => string.Join("|", error.Code, error.SheetName, error.RowIndex,
            error.ColumnIndex, error.PropertyName, error.FirstRowNumber, error.Message))
        .ToArray();

    /// <summary>
    /// 转换为关系合同数据。
    /// </summary>
    /// <param name="workbook">工作簿对象。</param>
    /// <returns>包含父级订单编号及其子项名称的可比较字符串列表。</returns>
    private static IReadOnlyList<string> ToRelationRows(RelationContractWorkbook workbook) => workbook.Parents
        .Select(parent => string.Join("|", parent.OrderNo,
            string.Join(",", parent.Items.Select(item => item.Name))))
        .ToArray();

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class ContractWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<ContractRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class RegistrationWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<RegistrationRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class RegistrationRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Count { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class RuleContractWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<RuleContractRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class RuleContractRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        [ExcelRequired]
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置状态。
        /// </summary>
        [ValueMapping("Enabled", 1)]
        public int Status { get; set; }

        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class RequiredContractWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<RequiredContractRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class RequiredContractRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        [ExcelRequired]
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class RelationContractWorkbook
    {
        /// <summary>
        /// 获取父项集合。
        /// </summary>
        public List<RelationContractParent> Parents { get; } = new();
        /// <summary>
        /// 获取子项集合。
        /// </summary>
        public List<RelationContractChild> Children { get; } = new();
    }

    /// <summary>
    /// 表示关系映射测试中的父项数据。
    /// </summary>
    private sealed class RelationContractParent
    {
        /// <summary>
        /// 获取或设置订单号。
        /// </summary>
        public string OrderNo { get; set; }
        /// <summary>
        /// 获取明细项集合。
        /// </summary>
        [ExcelIgnore]
        public List<RelationContractChild> Items { get; } = new();
    }

    /// <summary>
    /// 表示关系映射测试中的子项数据。
    /// </summary>
    private sealed class RelationContractChild
    {
        /// <summary>
        /// 获取或设置订单号。
        /// </summary>
        public string OrderNo { get; set; }
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 提供测试场景使用的值转换器。
    /// </summary>
    private sealed class ContractUpperConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "contract-upper";
        /// <summary>
        /// 获取导入上下文集合。
        /// </summary>
        public List<(int RowIndex, int ColumnIndex)> ImportContexts { get; } = new();

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            ImportContexts.Add((context.RowIndex, context.ColumnIndex));
            value = context.Value?.ToString()?.ToUpperInvariant();
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return true;
        }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class ContractRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Count { get; set; }
        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
        /// <summary>
        /// 获取或设置是否启用。
        /// </summary>
        public bool Enabled { get; set; }
        /// <summary>
        /// 获取或设置类型。
        /// </summary>
        public ContractKind Kind { get; set; }
        /// <summary>
        /// 获取或设置日期。
        /// </summary>
        public DateTime Date { get; set; }
        /// <summary>
        /// 获取或设置可选值。
        /// </summary>
        public int? Optional { get; set; }
        /// <summary>
        /// 获取或设置值集合。
        /// </summary>
        [DynamicColumn]
        public Dictionary<string, object> Values { get; set; } = new();
    }

    /// <summary>
    /// 表示 Provider 合同测试中的场景分支。
    /// </summary>
    private enum ContractKind
    {
        /// <summary>
        /// 表示数据处于活动状态。
        /// </summary>
        Active,
        /// <summary>
        /// 表示数据等待后续处理。
        /// </summary>
        Pending
    }
}
