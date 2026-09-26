using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests.Integration;

/// <summary>
/// ClosedXML 提供程序原生边界的合同测试。
/// </summary>
/// <remarks>跨提供程序的公共语义由 ProviderContract.Tests 覆盖。</remarks>
public sealed class ClosedXmlProviderContractTest
{
    /// <summary>
    /// 验证 NPOI 与 ClosedXML 对公共表头和数据行高度的写入结果。
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
    /// 验证列数和物理单元格限制在所有 XLSX Provider 中统一拒绝超限输入。
    /// </summary>
    [Fact]
    public void ColumnAndCellResourceContract_ShouldRejectBeforeDomAcrossAllXlsxProviders()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new ContractRow { Name = "A", Quantity = 1, Amount = 2m } }));
        using var generated = new System.IO.MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var bytes = generated.ToArray();
        var importers = new IExcelImporter[]
        {
            new NpoiExcelImporter(),
            new Bing.Offices.Imports.MiniExcelExcelImporter(),
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
                using var source = new System.IO.MemoryStream(bytes, writable: false);
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
    /// 获取公共成员签名中可见的类型。
    /// </summary>
    /// <param name="member">待检查的公共成员。</param>
    /// <returns>公共成员返回值、参数或属性中暴露的类型序列。</returns>
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
            yield return property.PropertyType;
        else if (member is FieldInfo field)
            yield return field.FieldType;
    }

    /// <summary>
    /// 资源边界测试使用的 Workbook。
    /// </summary>
    private sealed class ContractWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<ContractRow> Rows { get; } = new();
    }

    /// <summary>
    /// 资源和行高测试使用的一行数据。
    /// </summary>
    private sealed class ContractRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;
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
