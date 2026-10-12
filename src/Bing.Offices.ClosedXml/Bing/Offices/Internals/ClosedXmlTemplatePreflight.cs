using System.IO.Compression;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Internals;

namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// 在创建 ClosedXML DOM 前检查未由专用桥接路径接管的 XLSX 部件。
/// </summary>
internal static class ClosedXmlTemplatePreflight
{
    /// <summary>
    /// ClosedXML Provider 名称。
    /// </summary>
    private const string Provider = "ClosedXML";

    /// <summary>
    /// 检查模板 ZIP 中是否包含当前路径不承诺保留的部件。
    /// </summary>
    /// <param name="template">待检查的模板流。</param>
    /// <param name="operation">触发预检的操作类型。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="limits">模板 XML 资源限制；为空时使用默认限制。</param>
    /// <exception cref="BingOfficesUnsupportedFeatureException">
    /// 模板包含不支持保留的部件，或流不满足预检要求。
    /// </exception>
    /// <exception cref="BingOfficesResourceLimitException">模板 XML 超出资源限制。</exception>
    public static void Validate(Stream template, BingOfficesOperation operation = BingOfficesOperation.Export,
        CancellationToken cancellationToken = default, ExcelResourceLimits limits = null)
    {
        if (template == null)
            return;
        limits ??= new ExcelResourceLimits();
        limits.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        if (!template.CanRead || !template.CanSeek)
            throw Unsupported("ClosedXML 模板预检要求可读取且可定位的 XLSX 流。", null, operation);

        var originalPosition = template.Position;
        template.Position = 0;
        try
        {
            using var archive = new ZipArchive(template, ZipArchiveMode.Read, leaveOpen: true);
            var pivotParts = new HashSet<string>(ExcelPivotTemplateParts.Find(archive, limits, Provider,
                operation, cancellationToken), StringComparer.OrdinalIgnoreCase);
            foreach (var entry in archive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var name = entry.FullName.Replace('\\', '/');
                if (name.StartsWith("xl/charts/", StringComparison.OrdinalIgnoreCase))
                    throw Unsupported("模板包含 ClosedXML 第一版不承诺保留的 Chart 部件。", name, operation);
                if (pivotParts.Contains(name))
                    throw Unsupported("模板包含 ClosedXML 第一版不承诺保留的 PivotTable 部件。", name, operation);
                if (name.StartsWith("xl/media/", StringComparison.OrdinalIgnoreCase))
                    throw Unsupported("模板包含 ClosedXML 第一版不承诺保留的 Image 部件。", name, operation);
                if (name.StartsWith("xl/tables/", StringComparison.OrdinalIgnoreCase))
                    throw Unsupported("模板包含 ClosedXML 第一版不承诺保留的 Table 部件。", name, operation);
                if (name.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase)
                    || name.EndsWith("vbaProjectSignature.bin", StringComparison.OrdinalIgnoreCase))
                    throw Unsupported("模板包含 ClosedXML 第一版不承诺保留的 Macro 部件。", name, operation);
                if (name.StartsWith("xl/externalLinks/", StringComparison.OrdinalIgnoreCase))
                    throw Unsupported("模板包含 ClosedXML 第一版不承诺保留的 External Link 部件。", name, operation);
            }
        }
        catch (InvalidDataException exception)
        {
            if (operation == BingOfficesOperation.Import)
                throw new BingOfficesImportException("模板不是有效的 XLSX ZIP。", exception,
                    Provider, BingOfficesStage.Preflight);
            throw new BingOfficesExportException("模板不是有效的 XLSX ZIP。", exception,
                Provider, BingOfficesStage.Preflight);
        }
        finally
        {
            template.Position = originalPosition;
        }
    }

    /// <summary>
    /// 创建带模板部件信息的 Unsupported 异常。
    /// </summary>
    /// <param name="message">异常消息。</param>
    /// <param name="entry">命中的 ZIP 部件名称。</param>
    /// <param name="operation">触发预检的操作类型。</param>
    /// <returns>包含 Provider、操作和阶段信息的异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string message, string entry,
        BingOfficesOperation operation) =>
        new(entry == null ? message : $"{message} Entry: {entry}", provider: Provider,
            operation: operation, stage: BingOfficesStage.Preflight);
}
