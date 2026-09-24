using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests.Integration;

/// <summary>
/// 验证 ClosedXML 与现有 Provider 的公共基础语义和 API 隔离。
/// </summary>
public sealed class ClosedXmlProviderContractTest
{
    /// <summary>
    /// 验证基础标量往返结果在 NPOI、MiniExcel 和 ClosedXML 之间一致。
    /// </summary>
    [Fact]
    public async Task BasicScalarContract_ShouldMatchNpoiAndMiniExcel()
    {
        var rows = new[]
        {
            new ContractRow { Name = "A", Quantity = 2, Amount = 1.25m },
            new ContractRow { Name = "B", Quantity = 5, Amount = 9.50m }
        };
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows));
        var import = ExcelImport.Workbook<ContractWorkbook>(workbook =>
            workbook.Sheet<ContractRow>("Data", root => root.Rows));

        var npoi = await RoundTripAsync(new NpoiExcelExporter(), new NpoiExcelImporter(), export, import);
        var mini = await RoundTripAsync(new MiniExcelExcelExporter(), new MiniExcelExcelImporter(), export, import);
        var closed = await RoundTripAsync(new ClosedXmlExcelExporter(), new ClosedXmlExcelImporter(), export, import);

        Assert.Equal(Snapshot(npoi.Workbook.Rows), Snapshot(mini.Workbook.Rows));
        Assert.Equal(Snapshot(npoi.Workbook.Rows), Snapshot(closed.Workbook.Rows));
    }

    /// <summary>
    /// 验证动态列往返结果在三个 XLSX Provider 之间一致。
    /// </summary>
    [Fact]
    public async Task DynamicColumnContract_ShouldMatchAcrossProviders()
    {
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "Region",
            DataType = typeof(string)
        };
        var rows = new[]
        {
            new DynamicContractRow
            {
                Code = "A",
                Amount = 1.25m,
                Values = new Dictionary<string, object> { ["region"] = "east" }
            },
            new DynamicContractRow
            {
                Code = "B",
                Amount = 9.5m,
                Values = new Dictionary<string, object> { ["region"] = "west" }
            }
        };
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows,
            sheet => sheet.DynamicColumns(row => row.Values, new[] { definition })));
        var import = ExcelImport.Workbook<DynamicContractWorkbook>(workbook => workbook
            .Sheet<DynamicContractRow>("Data", root => root.Rows,
                sheet => sheet.DynamicColumns(row => row.Values, new[] { definition })));

        var npoi = await RoundTripAsync(new NpoiExcelExporter(), new NpoiExcelImporter(), export, import);
        var mini = await RoundTripAsync(new MiniExcelExcelExporter(), new MiniExcelExcelImporter(), export, import);
        var closed = await RoundTripAsync(new ClosedXmlExcelExporter(), new ClosedXmlExcelImporter(), export, import);

        Assert.True(npoi.IsSuccess, string.Join(";", npoi.Errors.Select(error => error.Message)));
        Assert.True(mini.IsSuccess, string.Join(";", mini.Errors.Select(error => error.Message)));
        Assert.True(closed.IsSuccess, string.Join(";", closed.Errors.Select(error => error.Message)));
        Assert.Equal(DynamicSnapshot(npoi.Workbook.Rows), DynamicSnapshot(mini.Workbook.Rows));
        Assert.Equal(DynamicSnapshot(npoi.Workbook.Rows), DynamicSnapshot(closed.Workbook.Rows));
    }

    /// <summary>
    /// 验证映射配置和显示值映射在三个 XLSX Provider 之间一致。
    /// </summary>
    [Fact]
    public async Task MappingContract_ShouldMatchAcrossAllXlsxProviders()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new()
                {
                    PropertyName = nameof(ContractRow.Name),
                    Title = "Display Name",
                    ValueMappings = new List<ExcelValueMappingConfiguration>
                    {
                        new() { Text = "Alice (display)", Value = "Alice" }
                    }
                },
                new() { PropertyName = nameof(ContractRow.Quantity), ColumnIndex = 2 }
            }
        };
        var rows = new[]
        {
            new ContractRow { Name = "Alice", Quantity = 2, Amount = 1.25m },
            new ContractRow { Name = "Bob", Quantity = 5, Amount = 9.50m }
        };
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows,
            sheet => sheet.Mapping(mapping)));
        var import = ExcelImport.Workbook<ContractWorkbook>(workbook => workbook
            .Sheet<ContractRow>("Data", root => root.Rows, sheet => sheet.Mapping(mapping)));

        var npoi = await RoundTripAsync(new NpoiExcelExporter(), new NpoiExcelImporter(), export, import);
        var mini = await RoundTripAsync(new MiniExcelExcelExporter(), new MiniExcelExcelImporter(), export, import);
        var closed = await RoundTripAsync(new ClosedXmlExcelExporter(), new ClosedXmlExcelImporter(), export, import);

        Assert.True(npoi.IsSuccess, string.Join(";", npoi.Errors.Select(error => error.Message)));
        Assert.True(mini.IsSuccess, string.Join(";", mini.Errors.Select(error => error.Message)));
        Assert.True(closed.IsSuccess, string.Join(";", closed.Errors.Select(error => error.Message)));
        Assert.Equal(Snapshot(npoi.Workbook.Rows), Snapshot(mini.Workbook.Rows));
        Assert.Equal(Snapshot(npoi.Workbook.Rows), Snapshot(closed.Workbook.Rows));
    }

    /// <summary>
    /// 验证公式单元格的缓存值和公式文本符合公共导入契约。
    /// </summary>
    [Fact]
    public void FormulaCachedValueContract_ShouldMatchNpoiAndClosedXml()
    {
        using var source = new MemoryStream();
        using (var workbook = new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("Formula");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue(nameof(FormulaContractRow.Amount));
            header.CreateCell(1).SetCellValue(nameof(FormulaContractRow.Enabled));
            header.CreateCell(2).SetCellValue(nameof(FormulaContractRow.Formula));
            var row = sheet.CreateRow(1);
            var amount = row.CreateCell(0);
            amount.SetCellFormula("1.5+2.25");
            var enabled = row.CreateCell(1);
            enabled.SetCellFormula("1=1");
            var formula = row.CreateCell(2);
            formula.SetCellFormula("\"Formula\"&\"Value\"");
            var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            evaluator.EvaluateFormulaCell(amount);
            evaluator.EvaluateFormulaCell(enabled);
            evaluator.EvaluateFormulaCell(formula);
            workbook.Write(source);
        }
        var formulaBytes = source.ToArray();
        var import = ExcelImport.Workbook<FormulaContractWorkbook>(workbook =>
            workbook.Sheet<FormulaContractRow>("Formula", root => root.Rows));

        using var npoiSource = new MemoryStream(formulaBytes, writable: false);
        using var closedSource = new MemoryStream(formulaBytes, writable: false);
        var npoi = new NpoiExcelImporter().Import(npoiSource, import);
        var closed = new ClosedXmlExcelImporter().Import(closedSource, import);

        Assert.True(npoi.IsSuccess, string.Join(";", npoi.Errors.Select(error => error.Message)));
        Assert.True(closed.IsSuccess, string.Join(";", closed.Errors.Select(error => error.Message)));
        Assert.Equal(FormulaSnapshot(npoi.Workbook.Rows), FormulaSnapshot(closed.Workbook.Rows));
        var npoiItem = Assert.Single(npoi.Workbook.Rows);
        var item = Assert.Single(closed.Workbook.Rows);
        Assert.Equal("FormulaValue", npoiItem.Formula);
        Assert.Equal(3.75m, item.Amount);
        Assert.True(item.Enabled);
        Assert.Equal("=\"Formula\"&\"Value\"", item.Formula);
    }

    /// <summary>
    /// 验证关系绑定及大小写不敏感比较在各 Provider 之间一致。
    /// </summary>
    [Fact]
    public async Task RelationContract_ShouldMatchAcrossProvidersIncludingClosedXml()
    {
        var export = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Parents", new[] { new RelationContractParent { OrderNo = "A-1" } })
            .AddSheet("Children", new[]
            {
                new RelationContractChild { OrderNo = "a-1", Name = "Item-1" },
                new RelationContractChild { OrderNo = "A-1", Name = "Item-2" }
            }));
        var import = ExcelImport.Workbook<RelationContractWorkbook>(workbook =>
        {
            workbook.Sheet("Parents", root => root.Parents);
            workbook.Sheet("Children", root => root.Children);
            workbook.HasMany(root => root.Parents, root => root.Children,
                parent => parent.OrderNo, child => child.OrderNo,
                parent => parent.Items, StringComparer.OrdinalIgnoreCase);
        });

        var npoi = await RoundTripAsync(new NpoiExcelExporter(), new NpoiExcelImporter(), export, import);
        var mini = await RoundTripAsync(new MiniExcelExcelExporter(), new MiniExcelExcelImporter(), export, import);
        var closed = await RoundTripAsync(new ClosedXmlExcelExporter(), new ClosedXmlExcelImporter(), export, import);

        Assert.True(npoi.IsSuccess, string.Join(";", npoi.Errors.Select(error => error.Message)));
        Assert.True(mini.IsSuccess, string.Join(";", mini.Errors.Select(error => error.Message)));
        Assert.True(closed.IsSuccess, string.Join(";", closed.Errors.Select(error => error.Message)));
        Assert.Equal(RelationSnapshot(npoi.Workbook), RelationSnapshot(mini.Workbook));
        Assert.Equal(RelationSnapshot(npoi.Workbook), RelationSnapshot(closed.Workbook));
    }

    /// <summary>
    /// 验证 NPOI 与 ClosedXML 对公共表头和数据行高度的写入结果一致。
    /// </summary>
    [Fact]
    public void RowHeightContract_ShouldMatchNpoiAndClosedXml()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new ContractRow { Name = "A", Quantity = 1, Amount = 1m } },
            sheet => sheet.RowHeight(new ExcelRowHeightOptions { HeaderHeight = 22, BodyHeight = 18 })));

        using var npoiStream = new System.IO.MemoryStream();
        new NpoiExcelExporter().Export(export, npoiStream);
        npoiStream.Position = 0;
        using var npoiWorkbook = new XSSFWorkbook(npoiStream);
        Assert.Equal(22, npoiWorkbook.GetSheet("Data").GetRow(0).HeightInPoints);
        Assert.Equal(18, npoiWorkbook.GetSheet("Data").GetRow(1).HeightInPoints);

        using var closedStream = new System.IO.MemoryStream();
        new ClosedXmlExcelExporter().Export(export, closedStream);
        closedStream.Position = 0;
        using var closedWorkbook = new XLWorkbook(closedStream);
        Assert.Equal(22, closedWorkbook.Worksheet("Data").Row(1).Height);
        Assert.Equal(18, closedWorkbook.Worksheet("Data").Row(2).Height);
    }

    /// <summary>
    /// 验证 NPOI 与 ClosedXML 在创建 DOM 前执行工作表资源限制。
    /// </summary>
    [Fact]
    public void SheetResourceLimitContract_ShouldRejectBeforeDomForNpoiAndClosedXml()
    {
        var export = ExcelExport.Workbook(workbook => workbook
            .AddSheet("One", new[] { new ContractRow { Name = "A" } })
            .AddSheet("Two", new[] { new ContractRow { Name = "B" } }));
        using var source = new System.IO.MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        var bytes = source.ToArray();
        var request = ExcelImport.Workbook<ContractWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxSheets = 1 })
            .Sheet<ContractRow>("One", root => root.Rows));

        using var npoiSource = new System.IO.MemoryStream(bytes, writable: false);
        Assert.Throws<BingOfficesResourceLimitException>(() =>
            new NpoiExcelImporter().Import(npoiSource, request));
        using var closedSource = new System.IO.MemoryStream(bytes, writable: false);
        Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(closedSource, request));
    }

    /// <summary>
    /// 验证校验失败结果在三个 XLSX Provider 之间保持一致。
    /// </summary>
    [Fact]
    public async Task ValidationContract_ShouldMatchAcrossAllXlsxProviders()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new RequiredContractRow { Code = string.Empty, Quantity = 0 } }));
        var import = ExcelImport.Workbook<RequiredContractWorkbook>(workbook => workbook
            .Sheet<RequiredContractRow>("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

        var npoi = await RoundTripAsync(new NpoiExcelExporter(), new NpoiExcelImporter(), export, import);
        var mini = await RoundTripAsync(new MiniExcelExcelExporter(), new MiniExcelExcelImporter(), export, import);
        var closed = await RoundTripAsync(new ClosedXmlExcelExporter(), new ClosedXmlExcelImporter(), export, import);

        Assert.False(npoi.IsSuccess);
        Assert.False(mini.IsSuccess);
        Assert.False(closed.IsSuccess);
        Assert.Contains(npoi.Errors, error => error.Code == ExcelImportErrorCode.Validation);
        Assert.Contains(mini.Errors, error => error.Code == ExcelImportErrorCode.Validation);
        Assert.Contains(closed.Errors, error => error.Code == ExcelImportErrorCode.Validation);
        Assert.Empty(npoi.Workbook.Rows);
        Assert.Empty(mini.Workbook.Rows);
        Assert.Empty(closed.Workbook.Rows);
    }

    /// <summary>
    /// 验证列数和物理单元格限制在所有 XLSX Provider 中统一拒绝超限输入。
    /// </summary>
    [Fact]
    public void ColumnAndCellResourceContract_ShouldRejectBeforeDomAcrossAllXlsxProviders()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new ContractRow { Name = "A", Quantity = 1, Amount = 2m } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var bytes = generated.ToArray();
        var importers = new IExcelImporter[]
        {
            new NpoiExcelImporter(),
            new MiniExcelExcelImporter(),
            new ClosedXmlExcelImporter()
        };
        foreach (var limit in new[]
        {
            new ExcelResourceLimits { MaxColumnsPerSheet = 2 },
            new ExcelResourceLimits { MaxCells = 3 }
        })
        {
            var request = ExcelImport.Workbook<ContractWorkbook>(workbook => workbook
                .ResourceLimits(limit)
                .Sheet<ContractRow>("Data", root => root.Rows));
            foreach (var importer in importers)
            {
                using var source = new MemoryStream(bytes, writable: false);
                Assert.Throws<BingOfficesResourceLimitException>(() =>
                    importer.Import(source, request));
            }
        }
    }

    /// <summary>
    /// 验证 ClosedXML 公共 Provider 类型不会暴露 ClosedXML DOM 类型。
    /// </summary>
    [Fact]
    public void ClosedXmlPublicProviderTypes_ShouldNotExposeClosedXmlDomTypes()
    {
        var assembly = typeof(ClosedXmlExcelExporter).Assembly;
        var publicTypes = assembly.GetExportedTypes()
            .Where(type => type.Namespace?.StartsWith("Bing.Offices.ClosedXml", StringComparison.Ordinal) == true)
            .ToArray();
        Assert.Contains(publicTypes, type => type == typeof(ClosedXmlExcelExporter));
        Assert.Contains(publicTypes, type => type == typeof(ClosedXmlExcelImporter));
        Assert.Contains(publicTypes, type => type == typeof(ExcelClosedXmlServiceCollectionExtensions));

        foreach (var type in publicTypes)
        {
            foreach (var member in type.GetMembers(BindingFlags.Public | BindingFlags.Instance
                | BindingFlags.Static | BindingFlags.DeclaredOnly))
            {
                var exposedTypes = GetExposedTypes(member).ToArray();
                Assert.DoesNotContain(exposedTypes, exposed =>
                    exposed.Namespace?.StartsWith("ClosedXML", StringComparison.Ordinal) == true);
            }
        }
    }

    /// <summary>
    /// 通过外围异步流完成导出和导入往返。
    /// </summary>
    /// <typeparam name="TWorkbook">导入工作簿模型类型。</typeparam>
    /// <param name="exporter">用于生成 XLSX 的导出器。</param>
    /// <param name="importer">用于读取 XLSX 的导入器。</param>
    /// <param name="export">导出请求。</param>
    /// <param name="import">导入请求。</param>
    /// <returns>导入后的工作簿结果。</returns>
    private static async Task<ExcelWorkbookImportResult<TWorkbook>> RoundTripAsync<TWorkbook>(
        IExcelExporter exporter, IExcelImporter importer, ExcelWorkbookExportRequest export,
        ExcelWorkbookImportRequest<TWorkbook> import)
        where TWorkbook : class, new()
    {
        await using var stream = new MemoryStream();
        await exporter.ExportAsync(export, stream);
        stream.Position = 0;
        return await importer.ImportAsync(stream, import);
    }

    /// <summary>
    /// 将标量数据行转换为稳定的契约快照。
    /// </summary>
    /// <param name="rows">待快照的数据行。</param>
    /// <returns>按输入顺序生成的行快照。</returns>
    private static IReadOnlyList<string> Snapshot(IEnumerable<ContractRow> rows) => rows
        .Select(row => $"{row.Name}|{row.Quantity}|{row.Amount.ToString(System.Globalization.CultureInfo.InvariantCulture)}")
        .ToArray();

    /// <summary>
    /// 将动态列数据行转换为稳定的契约快照。
    /// </summary>
    /// <param name="rows">待快照的数据行。</param>
    /// <returns>按输入顺序生成的动态列快照。</returns>
    private static IReadOnlyList<string> DynamicSnapshot(IEnumerable<DynamicContractRow> rows) => rows
        .Select(row => $"{row.Code}|{row.Amount.ToString("0.00", CultureInfo.InvariantCulture)}|"
            + (row.Values.TryGetValue("region", out var value) ? value?.ToString() : "<missing>"))
        .ToArray();

    /// <summary>
    /// 将公式数据行转换为稳定的缓存值快照。
    /// </summary>
    /// <param name="rows">待快照的数据行。</param>
    /// <returns>按输入顺序生成的公式快照。</returns>
    private static IReadOnlyList<string> FormulaSnapshot(IEnumerable<FormulaContractRow> rows) => rows
        .Select(row => $"{row.Amount.ToString(CultureInfo.InvariantCulture)}|{row.Enabled}")
        .ToArray();

    /// <summary>
    /// 将关系工作簿转换为父项及子项名称快照。
    /// </summary>
    /// <param name="workbook">待快照的关系工作簿。</param>
    /// <returns>按父项顺序生成的关系快照。</returns>
    private static IReadOnlyList<string> RelationSnapshot(RelationContractWorkbook workbook) =>
        workbook.Parents.Select(parent => $"{parent.OrderNo}|{string.Join(",",
            parent.Items.Select(item => item.Name))}").ToArray();

    /// <summary>
    /// 获取公共成员签名中可见的类型，供 API 隔离断言使用。
    /// </summary>
    /// <param name="member">待检查的反射成员。</param>
    /// <returns>成员返回值、参数或属性涉及的类型序列。</returns>
    private static IEnumerable<Type> GetExposedTypes(MemberInfo member)
    {
        if (member is MethodInfo method)
        {
            yield return method.ReturnType;
            foreach (var parameter in method.GetParameters())
                yield return parameter.ParameterType;
            foreach (var argument in method.GetGenericArguments())
                yield return argument;
        }
        else if (member is ConstructorInfo constructor)
        {
            foreach (var parameter in constructor.GetParameters())
                yield return parameter.ParameterType;
        }
        else if (member is PropertyInfo property)
        {
            yield return property.PropertyType;
        }
        else if (member is FieldInfo field)
        {
            yield return field.FieldType;
        }
    }

    /// <summary>
    /// 表示动态列契约测试的工作簿模型。
    /// </summary>
    private sealed class DynamicContractWorkbook
    {
        /// <summary>
        /// 获取动态列数据行集合。
        /// </summary>
        public List<DynamicContractRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示动态列契约测试的数据行。
    /// </summary>
    private sealed class DynamicContractRow
    {
        /// <summary>
        /// 获取或设置业务编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 获取或设置动态列值。
        /// </summary>
        [DynamicColumn]
        public Dictionary<string, object> Values { get; set; } = new();
    }

    /// <summary>
    /// 表示关系契约测试的工作簿模型。
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
    /// 表示关系契约测试中的父项。
    /// </summary>
    private sealed class RelationContractParent
    {
        /// <summary>
        /// 获取或设置父项订单号。
        /// </summary>
        public string OrderNo { get; set; }

        /// <summary>
        /// 获取关系绑定后的子项集合。
        /// </summary>
        [ExcelIgnore]
        public List<RelationContractChild> Items { get; } = new();
    }

    /// <summary>
    /// 表示关系契约测试中的子项。
    /// </summary>
    private sealed class RelationContractChild
    {
        /// <summary>
        /// 获取或设置父项订单号。
        /// </summary>
        public string OrderNo { get; set; }

        /// <summary>
        /// 获取或设置子项名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 表示基础标量契约测试的工作簿模型。
    /// </summary>
    private sealed class ContractWorkbook
    {
        /// <summary>
        /// 获取标量数据行集合。
        /// </summary>
        public List<ContractRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示校验契约测试的工作簿模型。
    /// </summary>
    private sealed class RequiredContractWorkbook
    {
        /// <summary>
        /// 获取待校验数据行集合。
        /// </summary>
        public List<RequiredContractRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示公式契约测试的工作簿模型。
    /// </summary>
    private sealed class FormulaContractWorkbook
    {
        /// <summary>
        /// 获取公式数据行集合。
        /// </summary>
        public List<FormulaContractRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示公式契约测试的数据行。
    /// </summary>
    private sealed class FormulaContractRow
    {
        /// <summary>
        /// 获取或设置金额缓存值。
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 获取或设置布尔缓存值。
        /// </summary>
        public bool Enabled { get; set; }

        /// <summary>
        /// 获取或设置公式文本。
        /// </summary>
        public string Formula { get; set; }
    }

    /// <summary>
    /// 表示必填校验契约测试的数据行。
    /// </summary>
    private sealed class RequiredContractRow
    {
        /// <summary>
        /// 获取或设置必填编码。
        /// </summary>
        [ExcelRequired]
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 表示基础标量契约测试的数据行。
    /// </summary>
    private sealed class ContractRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
    }
}
