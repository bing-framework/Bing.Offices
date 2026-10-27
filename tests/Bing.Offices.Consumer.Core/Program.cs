using System.Text;
using Bing.Offices.Attributes;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Microsoft.Extensions.DependencyInjection;

var services = new ServiceCollection();
services.AddBingOfficesCore();
using var provider = services.BuildServiceProvider();

Ensure(provider.GetService<IExcelImporter>() == null, "Core 包不应注册 Excel 导入 Provider。");
Ensure(provider.GetService<IExcelExporter>() == null, "Core 包不应注册 Excel 导出 Provider。");

var csvExporter = provider.GetRequiredService<ICsvExporter>();
var csvImporter = provider.GetRequiredService<ICsvImporter>();
var exportOptions = new CsvExportOptions<CoreConsumerRow> { Encoding = new UTF8Encoding(false, true) };
var importOptions = new CsvImportOptions<CoreConsumerRow> { Encoding = new UTF8Encoding(false, true) };
var expectedCsv = "Code\r\nORD-123\r\n";
var rows = new[] { new CoreConsumerRow { Code = "ORD-123" } };

var syncBytes = csvExporter.ExportToBytes(rows, exportOptions);
var asyncBytes = await csvExporter.ExportToBytesAsync(rows, exportOptions);
Ensure(Encoding.UTF8.GetString(syncBytes) == expectedCsv, "Core 同步 CSV 导出内容不完整或不正确。");
Ensure(Encoding.UTF8.GetString(asyncBytes) == expectedCsv, "Core 异步 CSV 导出内容不完整或不正确。");
Ensure(syncBytes.SequenceEqual(asyncBytes), "Core 同步与异步 CSV 导出内容不一致。");

var syncSuccess = csvImporter.ImportFromBytes<CoreConsumerRow>(syncBytes, importOptions);
var asyncSuccess = await csvImporter.ImportFromBytesAsync<CoreConsumerRow>(asyncBytes, importOptions);
AssertSuccess(syncSuccess);
AssertSuccess(asyncSuccess);

var invalidCsv = Encoding.UTF8.GetBytes("Code\r\ninvalid\r\nORD-123\r\n");
var syncFailure = csvImporter.ImportFromBytes<CoreConsumerRow>(invalidCsv, importOptions);
var asyncFailure = await csvImporter.ImportFromBytesAsync<CoreConsumerRow>(invalidCsv, importOptions);
AssertValidationFailure(syncFailure);
AssertValidationFailure(asyncFailure);

Console.WriteLine($"core-package-consumer-ok tfm={System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription} package=Bing.Offices.Core/{Environment.GetEnvironmentVariable("BING_OFFICES_PACKAGE_VERSION") ?? "2.0.0"} csv=sync-async excelProviders=absent");

static void AssertSuccess(CsvImportResult<CoreConsumerRow> result)
{
    Ensure(result.Items.Select(item => item.Code).SequenceEqual(new[] { "ORD-123" }),
        "Core CSV 导入未返回完整实体内容。");
    Ensure(result.Errors.Count == 0, "Core CSV 成功结果意外包含错误。");
    Ensure(result.IsSuccess, "Core CSV 成功结果的 IsSuccess 不正确。");
    Ensure(result.IsComplete, "Core CSV 成功结果的 IsComplete 不正确。");
    Ensure(result.StopReason == ImportStopReason.None, "Core CSV 成功结果的 StopReason 不正确。");
    Ensure(!result.IsTruncated && result.MaxErrors == null, "Core CSV 成功结果的截断状态不正确。");
}

static void AssertValidationFailure(CsvImportResult<CoreConsumerRow> result)
{
    Ensure(result.Items.Select(item => item.Code).SequenceEqual(new[] { "ORD-123" }),
        "Core CSV 错误结果未保留后续有效实体。");
    var error = result.Errors.Single();
    Ensure(error.Code == CsvImportErrorCode.Validation, "Core CSV 错误 Code 不正确。");
    Ensure(error.Message == "格式不正确", "Core CSV 错误 Message 不正确。");
    Ensure(error.RowIndex == 2 && error.ColumnIndex == 1, "Core CSV 错误行列位置不正确。");
    Ensure(error.PropertyName == nameof(CoreConsumerRow.Code), "Core CSV 错误属性名不正确。");
    Ensure(error.FirstRowNumber == null, "Core CSV 非重复错误不应包含 FirstRowNumber。");
    Ensure(!result.IsSuccess, "Core CSV 校验失败不应成功。");
    Ensure(result.IsComplete, "普通 CSV 行校验错误不应标记结果不完整。");
    Ensure(result.StopReason == ImportStopReason.None, "Core CSV 普通校验错误的 StopReason 不正确。");
    Ensure(!result.IsTruncated && result.MaxErrors == null, "Core CSV 校验错误的截断状态不正确。");
}

static void Ensure(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

public sealed class CoreConsumerRow
{
    [ExcelRegex(@"^ORD-\d+$")]
    public string Code { get; set; }
}
