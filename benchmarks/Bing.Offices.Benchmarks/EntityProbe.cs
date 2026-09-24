using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Bing.Offices.Entities;
using Bing.Offices.Extensions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 使用公开 NPOI Entity/Template API 的受控端到端探针。
/// </summary>
internal static class EntityProbe
{
    /// <summary>
    /// 执行固定单元格、合并区域、明细列表和模板 roundtrip。
    /// </summary>
    /// <param name="artifactPath">JSON 输出路径。</param>
    /// <param name="rowCount">每次实体明细行数。</param>
    /// <param name="repetitions">每种模式的测量次数。</param>
    /// <param name="phase">候选阶段标识，通常为 before 或 after。</param>
    public static async Task RunAsync(string artifactPath, int rowCount, int repetitions, string phase)
    {
        if (rowCount < 1)
            throw new ArgumentOutOfRangeException(nameof(rowCount));
        if (repetitions < 1)
            throw new ArgumentOutOfRangeException(nameof(repetitions));
        if (!string.Equals(phase, "before", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(phase, "after", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("阶段标识必须为 before 或 after。", nameof(phase));

        var fullPath = Path.GetFullPath(artifactPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var rows = Enumerable.Range(1, rowCount)
            .Select(index => new InvoiceLine { Code = $"L-{index:D6}", Quantity = index })
            .ToList();
        var entity = new Invoice { Number = 1001, Title = "Entity probe", Lines = rows };
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Merge("单据", "A1:B1")
            .Cell("单据", "B1", item => item.Title)
            .Cell("单据", "B2", item => item.Number)
            .ListRegion("明细", "A1", item => item.Lines));
        var templateBytes = CreateTemplate();
        using var services = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var exporter = services.GetRequiredService<IExcelExporter>();
        var importer = services.GetRequiredService<IExcelImporter>();
        var samples = new List<Sample>();

        for (var repetition = 1; repetition <= repetitions; repetition++)
        {
            samples.Add(Measure(exporter, importer, entity, layout, templateBytes,
                useTemplate: false, repetition));
            samples.Add(await MeasureAsync(exporter, importer, entity, layout, templateBytes,
                useTemplate: false, repetition).ConfigureAwait(false));
            samples.Add(Measure(exporter, importer, entity, layout, templateBytes,
                useTemplate: true, repetition));
            samples.Add(await MeasureAsync(exporter, importer, entity, layout, templateBytes,
                useTemplate: true, repetition).ConfigureAwait(false));
        }

        var document = new
        {
            kind = "npoi-entity-controlled-probe",
            schema = 1,
            generatedUtc = DateTimeOffset.UtcNow,
            framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            rowCount,
            repetitions,
            phase = phase.ToLowerInvariant(),
            workload = "fixed-cell-merge-list-template-roundtrip",
            candidateIdentity = GetCandidateIdentity(),
            samples
        };
        await File.WriteAllTextAsync(fullPath, JsonSerializer.Serialize(document), new UTF8Encoding(false))
            .ConfigureAwait(false);
        Console.WriteLine($"ENTITY_PROBE artifact={fullPath} rows={rowCount} repetitions={repetitions} status=passed");
    }

    /// <summary>
    /// 同步测量一次实体导出和导入往返。
    /// </summary>
    /// <typeparam name="TEntity">待测实体类型。</typeparam>
    /// <param name="exporter">实体导出器。</param>
    /// <param name="importer">实体导入器。</param>
    /// <param name="entity">待导出的实体。</param>
    /// <param name="layout">实体布局定义。</param>
    /// <param name="templateBytes">模板工作簿字节。</param>
    /// <param name="useTemplate">是否使用模板路径。</param>
    /// <param name="repetition">当前测量重复序号。</param>
    /// <returns>本次同步测量结果。</returns>
    private static Sample Measure<TEntity>(IExcelExporter exporter, IExcelImporter importer,
        TEntity entity, ExcelEntityLayout<TEntity> layout, byte[] templateBytes, bool useTemplate,
        int repetition) where TEntity : class, new()
    {
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        var allocatedBefore = GC.GetTotalAllocatedBytes(true);
        var stopwatch = Stopwatch.StartNew();
        using var output = new MemoryStream();
        if (useTemplate)
        {
            using var template = new MemoryStream(templateBytes, writable: false);
            exporter.ExportForTemplate(entity, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        }
        else
        {
            exporter.ExportEntity(entity, layout, output);
        }

        var outputBytes = output.ToArray();
        using var source = new MemoryStream(outputBytes, writable: false);
        ExcelEntityImportResult<TEntity> result;
        if (useTemplate)
        {
            using var template = new MemoryStream(templateBytes, writable: false);
            result = importer.ImportForTemplate(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true));
        }
        else
        {
            result = importer.ImportEntity(source, layout);
        }

        stopwatch.Stop();
        Ensure(result.IsSuccess, "Entity probe returned structured errors.");
        var imported = (Invoice)(object)result.Entity;
        Ensure(imported.Number == ((Invoice)(object)entity).Number
            && imported.Lines.Count == ((Invoice)(object)entity).Lines.Count,
            "Entity probe roundtrip count mismatch.");
        if (useTemplate)
            VerifyTemplateOutput(outputBytes, imported.Lines.Count);
        return new Sample(useTemplate ? "template-sync" : "entity-sync", repetition,
            stopwatch.Elapsed.TotalMilliseconds, GC.GetTotalAllocatedBytes(true) - allocatedBefore,
            outputBytes.Length, imported.Lines.Count);
    }

    /// <summary>
    /// 异步测量一次实体导出和导入往返。
    /// </summary>
    /// <typeparam name="TEntity">待测实体类型。</typeparam>
    /// <param name="exporter">实体导出器。</param>
    /// <param name="importer">实体导入器。</param>
    /// <param name="entity">待导出的实体。</param>
    /// <param name="layout">实体布局定义。</param>
    /// <param name="templateBytes">模板工作簿字节。</param>
    /// <param name="useTemplate">是否使用模板路径。</param>
    /// <param name="repetition">当前测量重复序号。</param>
    /// <returns>最终完成的异步测量结果。</returns>
    private static async Task<Sample> MeasureAsync<TEntity>(IExcelExporter exporter, IExcelImporter importer,
        TEntity entity, ExcelEntityLayout<TEntity> layout, byte[] templateBytes, bool useTemplate,
        int repetition) where TEntity : class, new()
    {
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        var allocatedBefore = GC.GetTotalAllocatedBytes(true);
        var stopwatch = Stopwatch.StartNew();
        using var output = new MemoryStream();
        if (useTemplate)
        {
            using var template = new MemoryStream(templateBytes, writable: false);
            await exporter.ExportForTemplateAsync(entity, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output).ConfigureAwait(false);
        }
        else
        {
            await exporter.ExportEntityAsync(entity, layout, output).ConfigureAwait(false);
        }

        var outputBytes = output.ToArray();
        using var source = new MemoryStream(outputBytes, writable: false);
        ExcelEntityImportResult<TEntity> result;
        if (useTemplate)
        {
            using var template = new MemoryStream(templateBytes, writable: false);
            result = await importer.ImportForTemplateAsync(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true)).ConfigureAwait(false);
        }
        else
        {
            result = await importer.ImportEntityAsync(source, layout).ConfigureAwait(false);
        }

        stopwatch.Stop();
        Ensure(result.IsSuccess, "Entity probe returned structured errors.");
        var imported = (Invoice)(object)result.Entity;
        Ensure(imported.Number == ((Invoice)(object)entity).Number
            && imported.Lines.Count == ((Invoice)(object)entity).Lines.Count,
            "Entity probe roundtrip count mismatch.");
        if (useTemplate)
            VerifyTemplateOutput(outputBytes, imported.Lines.Count);
        return new Sample(useTemplate ? "template-async" : "entity-async", repetition,
            stopwatch.Elapsed.TotalMilliseconds, GC.GetTotalAllocatedBytes(true) - allocatedBefore,
            outputBytes.Length, imported.Lines.Count);
    }

    /// <summary>
    /// 校验模板导出结果中的关键结构。
    /// </summary>
    /// <param name="bytes">待校验的 XLSX 字节。</param>
    /// <param name="rowCount">预期的明细行数。</param>
    private static void VerifyTemplateOutput(byte[] bytes, int rowCount)
    {
        using var source = new MemoryStream(bytes, writable: false);
        using var workbook = WorkbookFactory.Create(source);
        var invoice = workbook.GetSheet("单据");
        Ensure(invoice.NumMergedRegions == 1, "Entity template merge was not preserved.");
        Ensure(invoice.GetRow(0).GetCell(2).CellFormula == "1+1", "Entity template formula was not preserved.");
        Ensure(workbook.GetAllPictures().Count == 1, "Entity template picture was not preserved.");
        Ensure(workbook.GetSheet("明细").LastRowNum == rowCount, "Entity template list row count mismatch.");
    }

    /// <summary>
    /// 创建实体探针使用的 XLSX 模板。
    /// </summary>
    /// <returns>模板工作簿字节。</returns>
    private static byte[] CreateTemplate()
    {
        using var workbook = new XSSFWorkbook();
        var invoice = workbook.CreateSheet("单据");
        var title = invoice.CreateRow(0).CreateCell(0);
        title.SetCellValue("Template title");
        invoice.GetRow(0).CreateCell(2).CellFormula = "1+1";
        invoice.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        workbook.CreateSheet("明细");
        var pictureIndex = workbook.AddPicture(TemplatePng, PictureType.PNG);
        var anchor = workbook.GetCreationHelper().CreateClientAnchor();
        anchor.Row1 = 2;
        anchor.Row2 = 4;
        anchor.Col1 = 2;
        anchor.Col2 = 4;
        invoice.CreateDrawingPatriarch().CreatePicture(anchor, pictureIndex);
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    /// <summary>
    /// 获取参与探针运行的程序集身份哈希。
    /// </summary>
    /// <returns>按文件名索引的程序集 SHA-256 哈希。</returns>
    private static IReadOnlyDictionary<string, string> GetCandidateIdentity()
    {
        var types = new[] { typeof(EntityProbe), typeof(IExcelExporter), typeof(Bing.Offices.Exports.NpoiExcelExporter) };
        return types.Select(type => type.Assembly.Location).Distinct(StringComparer.OrdinalIgnoreCase)
            .ToDictionary(path => Path.GetFileName(path), path => Convert.ToHexString(
                SHA256.HashData(File.ReadAllBytes(path))), StringComparer.Ordinal);
    }

    /// <summary>
    /// 在探针自检失败时抛出异常。
    /// </summary>
    /// <param name="condition">需要满足的条件。</param>
    /// <param name="message">条件不满足时使用的异常消息。</param>
    private static void Ensure(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }

    /// <summary>
    /// 模板中使用的最小 PNG 图片字节。
    /// </summary>
    private static readonly byte[] TemplatePng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    /// <summary>
    /// 记录一次实体探针测量结果。
    /// </summary>
    /// <param name="mode">测量模式。</param>
    /// <param name="repetition">测量重复序号。</param>
    /// <param name="elapsedMilliseconds">耗时，单位为毫秒。</param>
    /// <param name="allocatedBytes">托管堆分配字节数。</param>
    /// <param name="outputBytes">输出工作簿字节数。</param>
    /// <param name="importedRows">导入的明细行数。</param>
    private sealed record Sample(string mode, int repetition, double elapsedMilliseconds,
        long allocatedBytes, int outputBytes, int importedRows);

    /// <summary>
    /// 实体探针使用的发票实体。
    /// </summary>
    private sealed class Invoice
    {
        /// <summary>
        /// 获取或设置单据编号。
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// 获取或设置单据标题。
        /// </summary>
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// 获取或设置明细行集合。
        /// </summary>
        public List<InvoiceLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 实体探针使用的发票明细行。
    /// </summary>
    private sealed class InvoiceLine
    {
        /// <summary>
        /// 获取或设置明细编码。
        /// </summary>
        public string Code { get; set; } = string.Empty;

        /// <summary>
        /// 获取或设置明细数量。
        /// </summary>
        public int Quantity { get; set; }
    }
}
