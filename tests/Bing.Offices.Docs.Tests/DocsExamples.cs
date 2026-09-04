using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Extensions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Docs.Tests;

internal static class DocsExamples
{
    internal static void ReadmeExport(IExcelExporter exporter, Stream stream, IEnumerable<OrderExport> orders)
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("订单", orders));
        exporter.Export(request, stream);
    }

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

    internal static ExcelMappingDocument JsonXml(string json)
    {
        var document = ExcelMappingConfigurationLoader.FromJsonDocument(json);
        var importRequest = ExcelImport.Workbook<OrdersWorkbook>(builder =>
            builder.Sheet("订单", workbook => workbook.Items, sheet => sheet.Mapping(document)));
        _ = importRequest;
        var migrated = ExcelMappingConfigurationLoader.MigrateV1Json(
            "{\"columns\":[{\"propertyName\":\"Code\",\"title\":\"编码\"}]}",
            MappingDirection.Import, out var diagnostics);
        if (!diagnostics.Any(item => item.Code == "V1_MIGRATED"))
            throw new InvalidOperationException("v1 migration diagnostic missing");
        return migrated;
    }

    internal static CsvImportResult<ValidatedRow> Validation(Stream source, ICsvImporter importer)
    {
        var result = importer.Import<ValidatedRow>(source);
        return result;
    }

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
