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

internal static class DocsExamples
{
    // docgen:sample name=ExportWorkbook title=导出一个工作簿
    internal static void ReadmeExport(IExcelExporter exporter, Stream stream, IEnumerable<OrderExport> orders)
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("订单", orders));
        exporter.Export(request, stream);
    }

    // docgen:sample name=MappingProfile title=用 Profile 声明导入导出映射
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

    // docgen:sample name=JsonXmlMapping title=用 JSON / XML 映射文档驱动导入
    internal static ExcelMappingDocument JsonXml(string json)
    {
        var document = ExcelMappingConfigurationLoader.FromJsonDocument(json);
        var importRequest = ExcelImport.Workbook<OrdersWorkbook>(builder =>
            builder.Sheet("订单", workbook => workbook.Items, sheet => sheet.Mapping(document)));
        _ = importRequest;
        return document;
    }

    // docgen:sample name=CsvValidation title=CSV 导入并触发实体验证
    internal static CsvImportResult<ValidatedRow> Validation(Stream source, ICsvImporter importer)
    {
        var result = importer.Import<ValidatedRow>(source);
        return result;
    }

    // docgen:sample name=AspnetCoreUpload title=ASP.NET Core 上传接口中导入工作簿
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

    // docgen:sample name=DynamicColumns title=导入动态列到字典
    internal static CsvImportResult<DynamicRow> Dynamic(Stream input, ICsvImporter importer)
    {
        var result = importer.Import<DynamicRow>(input);
        return result;
    }

    internal sealed class OrderProfile : IMappingProfile<OrderImport, OrderExport>
    {
        public void Configure(FluentSetting<OrderImport, OrderExport> setting)
        {
            setting.Import.Property(order => order.Code).HasHeader("订单号").HasAlias("旧订单号");
            setting.Export.Property(order => order.DisplayName).HasHeader("客户名称").HasFormatter("@");
        }
    }

    internal sealed class OrderImport
    {
        public string Code { get; set; }
    }

    internal sealed class OrderExport
    {
        public string DisplayName { get; set; }
    }

    internal sealed class OrdersWorkbook
    {
        public ICollection<OrderImport> Items { get; } = new List<OrderImport>();
    }

    internal sealed class UploadWorkbook
    {
        public ICollection<ValidatedRow> Rows { get; } = new List<ValidatedRow>();
    }

    internal sealed class ValidatedRow
    {
        [ExcelRequired]
        [ExcelRegex("^ORD-")]
        [ExcelUnique]
        public string Code { get; set; }

        [ExcelDate("yyyy-MM-dd")]
        public DateTime CreatedAt { get; set; }
    }

    internal sealed class DynamicRow
    {
        public string Name { get; set; }
        [DynamicColumn]
        public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>();
    }
}
