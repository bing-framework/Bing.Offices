using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Internals;

/// <summary>
/// 保存透视模板的磁盘快照及执行计划。
/// </summary>
internal sealed class ExcelPivotTemplateSnapshot : IDisposable
{
    /// <summary>
    /// 由当前快照持有的临时文件流。
    /// </summary>
    private readonly FileStream _source;

    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplateSnapshot" /> 类型的实例。
    /// </summary>
    /// <param name="source">已完成预检的临时文件。</param>
    /// <param name="plan">不可变执行计划。</param>
    /// <param name="limits">快照使用的资源限制。</param>
    private ExcelPivotTemplateSnapshot(FileStream source, ExcelPivotTemplatePlan plan,
        ExcelResourceLimits limits)
    {
        _source = source;
        Plan = plan;
        Limits = limits;
    }

    /// <summary>
    /// 获取模板快照流。
    /// </summary>
    internal Stream Source => _source;

    /// <summary>
    /// 获取执行计划。
    /// </summary>
    internal ExcelPivotTemplatePlan Plan { get; }

    /// <summary>
    /// 获取快照使用的资源限制。
    /// </summary>
    internal ExcelResourceLimits Limits { get; }

    /// <summary>
    /// 创建透视模板快照。
    /// </summary>
    /// <param name="request">包含模板的工作簿请求。</param>
    /// <param name="limits">输入及 ZIP/XML 资源限制。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>持有临时文件的快照。</returns>
    /// <remarks>不接管输入流；成功、失败或取消均恢复输入位置，临时文件随快照释放。</remarks>
    internal static ExcelPivotTemplateSnapshot Create(ExcelWorkbookExportRequest request,
        ExcelResourceLimits limits, string provider, CancellationToken token = default)
    {
        var input = Validate(request, limits, provider, token);
        var position = input.Position;
        var staged = Open();
        try
        {
            input.Position = 0;
            var buffer = new byte[81920];
            long total = 0;
            int count;
            while ((count = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                token.ThrowIfCancellationRequested();
                CheckLength(total, count, limits, provider);
                staged.Write(buffer, 0, count);
                total += count;
            }
            return Complete(request, limits, provider, staged, token);
        }
        catch
        {
            staged.Dispose();
            throw;
        }
        finally
        {
            input.Position = position;
        }
    }

    /// <summary>
    /// 创建透视模板快照。
    /// </summary>
    /// <param name="request">包含模板的工作簿请求。</param>
    /// <param name="limits">输入及 ZIP/XML 资源限制。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>持有临时文件的快照。</returns>
    /// <remarks>外围输入复制使用真实异步 IO；XML 预检和计划生成沿用相同执行语义。</remarks>
    internal static async Task<ExcelPivotTemplateSnapshot> CreateAsync(ExcelWorkbookExportRequest request,
        ExcelResourceLimits limits, string provider, CancellationToken token = default)
    {
        var input = Validate(request, limits, provider, token);
        var position = input.Position;
        var staged = Open();
        try
        {
            input.Position = 0;
            var buffer = new byte[81920];
            long total = 0;
            int count;
            while ((count = await input.ReadAsync(buffer.AsMemory(), token).ConfigureAwait(false)) > 0)
            {
                CheckLength(total, count, limits, provider);
                await staged.WriteAsync(buffer.AsMemory(0, count), token).ConfigureAwait(false);
                total += count;
            }
            await staged.FlushAsync(token).ConfigureAwait(false);
            return Complete(request, limits, provider, staged, token);
        }
        catch
        {
            await staged.DisposeAsync().ConfigureAwait(false);
            throw;
        }
        finally
        {
            input.Position = position;
        }
    }

    /// <inheritdoc />
    public void Dispose() => _source.Dispose();

    /// <summary>
    /// 校验快照输入。
    /// </summary>
    /// <param name="request">工作簿请求。</param>
    /// <param name="limits">资源限制。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>调用方持有的模板流。</returns>
    private static Stream Validate(ExcelWorkbookExportRequest request, ExcelResourceLimits limits,
        string provider, CancellationToken token)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (limits == null)
            throw new ArgumentNullException(nameof(limits));
        token.ThrowIfCancellationRequested();
        limits.Validate();
        var input = request.Template;
        if (input == null || !input.CanRead || !input.CanSeek)
            throw new BingOfficesUnsupportedFeatureException("透视模板快照要求可读取且可定位的模板流。",
                provider: provider, operation: BingOfficesOperation.Export, stage: BingOfficesStage.Preflight);
        CheckLength(0, input.Length, limits, provider);
        return input;
    }

    /// <summary>
    /// 创建随关闭清理的临时文件。
    /// </summary>
    /// <returns>可读写的异步临时文件流。</returns>
    private static FileStream Open() => new(Path.Combine(Path.GetTempPath(),
        "bing-offices-pivot-template-" + Guid.NewGuid().ToString("N") + ".tmp"),
        FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920,
        FileOptions.DeleteOnClose | FileOptions.Asynchronous);

    /// <summary>
    /// 完成模板资源预检及更新定位。
    /// </summary>
    /// <param name="request">工作簿请求。</param>
    /// <param name="limits">资源限制。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="staged">完整临时文件。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>通过预检的快照。</returns>
    private static ExcelPivotTemplateSnapshot Complete(ExcelWorkbookExportRequest request,
        ExcelResourceLimits limits, string provider, FileStream staged, CancellationToken token)
    {
        try
        {
            staged.Position = 0;
            var manifest = ExcelPivotTemplateManifest.Read(staged, limits, provider, token);
            if (request.PivotTableUpdates.Count > 0 && manifest.Entries.Count == 0)
                throw new BingOfficesUnsupportedFeatureException(
                    $"{provider}: 模板不包含可更新的透视表，当前 Provider 不支持该更新。", provider: provider,
                    operation: BingOfficesOperation.Export, stage: BingOfficesStage.Preflight);
            var plan = ExcelPivotTemplatePlan.Create(request, manifest, provider, token);
            staged.Position = 0;
            return new ExcelPivotTemplateSnapshot(staged, plan, limits);
        }
        catch (BingOfficesResourceLimitException error)
        {
            throw new BingOfficesResourceLimitException(error.Message, error, provider,
                BingOfficesOperation.Export, BingOfficesStage.Preflight);
        }
        catch (Exception error) when (error is InvalidDataException or BingOfficesImportException)
        {
            throw new BingOfficesExportException("透视模板预检失败。", error, provider, BingOfficesStage.Preflight);
        }
    }

    /// <summary>
    /// 检查快照输入字节预算。
    /// </summary>
    /// <param name="total">已读取字节数。</param>
    /// <param name="count">本次读取字节数。</param>
    /// <param name="limits">资源限制。</param>
    /// <param name="provider">提供程序名称。</param>
    private static void CheckLength(long total, long count, ExcelResourceLimits limits, string provider)
    {
        if (limits.MaxInputBytes.HasValue && count > limits.MaxInputBytes.Value - total)
            throw new BingOfficesResourceLimitException("透视模板输入大小超过限制。", provider: provider,
                operation: BingOfficesOperation.Export);
    }
}
