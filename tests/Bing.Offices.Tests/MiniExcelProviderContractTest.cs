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
using Bing.Offices.Npoi.Extensions;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI 与 MiniExcel 共同 XLSX 能力的结构化合同测试。
/// </summary>
public sealed class MiniExcelProviderContractTest
{
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
                .Mapping(mapping).Validate(ValidateMode.Continue)));

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

    [Fact]
    public async Task StructuredErrorContract_ShouldMatchAcrossProviders()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new RequiredContractRow { Code = string.Empty, Quantity = 0 } }));
        var importRequest = ExcelImport.Workbook<RequiredContractWorkbook>(workbook => workbook
            .Sheet<RequiredContractRow>("Data", root => root.Rows,
                sheet => sheet.Validate(ValidateMode.Continue)));

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
        Assert.Throws<Bing.Offices.Exceptions.BingOfficesUnsupportedFeatureException>(() =>
            new MiniExcelExcelExporter().Export(request, miniOutput));
        Assert.Equal(0, miniOutput.Length);
        Assert.True(npoiOutput.Length > 0);
    }

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

    private static IReadOnlyList<string> ToContractRows(IReadOnlyList<ContractRow> rows) => rows
        .Select(row => string.Join("|", row.Code, row.Count.ToString(CultureInfo.InvariantCulture),
            row.Amount.ToString("0.00", CultureInfo.InvariantCulture), row.Enabled,
            row.Kind, row.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            row.Optional?.ToString(CultureInfo.InvariantCulture) ?? "<null>",
            row.Values.TryGetValue("region", out var region) ? region?.ToString() : "<missing>"))
        .ToArray();

    private static IReadOnlyList<string> ToRuleRows(IReadOnlyList<RuleContractRow> rows) => rows
        .Select(row => string.Join("|", row.Code, row.Status,
            row.Name, row.Name == row.Name.ToUpperInvariant()))
        .ToArray();

    private static IReadOnlyList<string> ToErrorContract(IReadOnlyList<ExcelImportError> errors) => errors
        .Select(error => string.Join("|", error.Code, error.SheetName, error.RowIndex,
            error.ColumnIndex, error.PropertyName, error.FirstRowNumber, error.Message))
        .ToArray();

    private static IReadOnlyList<string> ToRelationRows(RelationContractWorkbook workbook) => workbook.Parents
        .Select(parent => string.Join("|", parent.OrderNo,
            string.Join(",", parent.Items.Select(item => item.Name))))
        .ToArray();

    private sealed class ContractWorkbook
    {
        public List<ContractRow> Rows { get; } = new();
    }

    private sealed class RuleContractWorkbook
    {
        public List<RuleContractRow> Rows { get; } = new();
    }

    private sealed class RuleContractRow
    {
        [ExcelRequired]
        public string Code { get; set; }

        [ValueMapping("Enabled", 1)]
        public int Status { get; set; }

        public string Name { get; set; }
    }

    private sealed class RequiredContractWorkbook
    {
        public List<RequiredContractRow> Rows { get; } = new();
    }

    private sealed class RequiredContractRow
    {
        [ExcelRequired]
        public string Code { get; set; }

        public int Quantity { get; set; }
    }

    private sealed class RelationContractWorkbook
    {
        public List<RelationContractParent> Parents { get; } = new();
        public List<RelationContractChild> Children { get; } = new();
    }

    private sealed class RelationContractParent
    {
        public string OrderNo { get; set; }
        [ExcelIgnore]
        public List<RelationContractChild> Items { get; } = new();
    }

    private sealed class RelationContractChild
    {
        public string OrderNo { get; set; }
        public string Name { get; set; }
    }

    private sealed class ContractUpperConverter : INamedExcelValueConverter
    {
        public string Name => "contract-upper";
        public List<(int RowIndex, int ColumnIndex)> ImportContexts { get; } = new();

        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            ImportContexts.Add((context.RowIndex, context.ColumnIndex));
            value = context.Value?.ToString()?.ToUpperInvariant();
            return true;
        }

        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return true;
        }
    }

    private sealed class ContractRow
    {
        public string Code { get; set; }
        public int Count { get; set; }
        public decimal Amount { get; set; }
        public bool Enabled { get; set; }
        public ContractKind Kind { get; set; }
        public DateTime Date { get; set; }
        public int? Optional { get; set; }
        [DynamicColumn]
        public Dictionary<string, object> Values { get; set; } = new();
    }

    private enum ContractKind
    {
        Active,
        Pending
    }
}
