using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Extensions;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Docs.Tests;

/// <summary>
/// 提供文档示例中的可执行导入、导出和校验调用。
/// </summary>
internal static class DocsExamples
{
    /// <summary>
    /// 执行 README 中的 Excel 导出示例。
    /// </summary>
    /// <param name="exporter">Excel 或 CSV 导出器。</param>
    /// <param name="stream">参与操作的流。</param>
    /// <param name="orders">订单集合。</param>
    internal static void ReadmeExport(IExcelExporter exporter, Stream stream, IEnumerable<OrderExport> orders)
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("订单", orders));
        exporter.Export(request, stream);
    }

    /// <summary>
    /// 创建并解析映射 Profile 描述。
    /// </summary>
    /// <returns>注册的订单 Profile 对应的导入映射描述。</returns>
    internal static ProfileDescriptor Profile()
    {
        var services = new ServiceCollection();
        services.AddMappingProfile<OrderProfile>();
        using var provider = services.BuildServiceProvider();
        var registry = provider.GetRequiredService<IMappingProfileRegistry>();
        if (!registry.TryGetDescriptor(typeof(OrderProfile).FullName, MappingDirection.Import,
                typeof(OrderImport), out var descriptor))
            throw new InvalidOperationException("Profile descriptor missing");
        return descriptor;
    }

    /// <summary>
    /// 根据 JSON 映射文本创建映射文档。
    /// </summary>
    /// <param name="json">JSON 映射文本。</param>
    /// <returns>从 JSON 解析的映射文档。</returns>
    internal static ExcelMappingDocument JsonXml(string json)
    {
        var document = ExcelMappingConfigurationLoader.FromJsonDocument(json);
        var importRequest = ExcelImport.Workbook<OrdersWorkbook>(builder =>
            builder.Sheet("订单", workbook => workbook.Items, sheet => sheet.Mapping(document)));
        _ = importRequest;
        return document;
    }

    /// <summary>
    /// 创建包含验证结果的 CSV 导入示例。
    /// </summary>
    /// <param name="source">输入流。</param>
    /// <param name="importer">Excel 或 CSV 导入器。</param>
    /// <returns>包含实体数据和校验错误的 CSV 导入结果。</returns>
    internal static CsvImportResult<ValidatedRow> Validation(Stream source, ICsvImporter importer)
    {
        var result = importer.Import<ValidatedRow>(source);
        return result;
    }

    /// <summary>
    /// 执行上传示例并返回 HTTP 结果。
    /// </summary>
    /// <param name="file">文件对象或路径。</param>
    /// <param name="importer">Excel 或 CSV 导入器。</param>
    /// <returns>成功时返回包含导入结果的 OK 响应，发生异常时返回 BadRequest 响应。</returns>
    internal static IResult Upload(IFormFile file, IExcelImporter importer)
    {
        try
        {
            using var input = file.OpenReadStream();
            var request = ExcelImport.Workbook<UploadWorkbook>(builder =>
                builder.Sheet("Data", workbook => workbook.Rows));
            var result = importer.Import(input, request);
            return Results.Ok(result);
        }
        catch (Exception)
        {
            return Results.BadRequest();
        }
    }

    /// <summary>
    /// 执行动态列 CSV 导入示例。
    /// </summary>
    /// <param name="input">输入数据。</param>
    /// <param name="importer">Excel 或 CSV 导入器。</param>
    /// <returns>包含动态列数据的 CSV 导入结果。</returns>
    internal static CsvImportResult<DynamicRow> Dynamic(Stream input, ICsvImporter importer)
    {
        var result = importer.Import<DynamicRow>(input);
        return result;
    }

    /// <summary>
    /// 提供测试场景使用的映射 Profile。
    /// </summary>
    internal sealed class OrderProfile : IMappingProfile<OrderImport, OrderExport>
    {
        /// <inheritdoc />
        public void Configure(FluentSetting<OrderImport, OrderExport> setting)
        {
            setting.Import.Property(order => order.Code).HasHeader("订单号").HasAlias("旧订单号");
            setting.Export.Property(order => order.DisplayName).HasHeader("客户名称").HasFormatter("@");
        }
    }

    /// <summary>
    /// 表示文档示例使用的导入订单。
    /// </summary>
    internal sealed class OrderImport
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; }
    }

    /// <summary>
    /// 表示文档示例使用的导出订单。
    /// </summary>
    internal sealed class OrderExport
    {
        /// <summary>
        /// 获取或设置显示名称。
        /// </summary>
        public string DisplayName { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    internal sealed class OrdersWorkbook
    {
        /// <summary>
        /// 获取明细项集合。
        /// </summary>
        public ICollection<OrderImport> Items { get; } = new List<OrderImport>();
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    internal sealed class UploadWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public ICollection<ValidatedRow> Rows { get; } = new List<ValidatedRow>();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    internal sealed class ValidatedRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        [ExcelRequired]
        [ExcelRegex("^ORD-")]
        [ExcelUnique]
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置创建时间。
        /// </summary>
        [ExcelDate("yyyy-MM-dd")]
        public DateTime CreatedAt { get; set; }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    internal sealed class DynamicRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置值集合。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>();
    }
}
