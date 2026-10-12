using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Internals;

/// <summary>
/// 复制仅保留模板的透视工作簿。
/// </summary>
internal static class ExcelPivotTemplatePreserver
{
    /// <summary>
    /// 判断请求是否启用模板透视保留模式。
    /// </summary>
    /// <param name="request">工作簿导出请求。</param>
    /// <returns>是否为模板透视保留模式。</returns>
    internal static bool IsPreserveTemplate(ExcelWorkbookExportRequest request) =>
        request?.PreserveTemplatePivotTables == true
        && request.Template != null
        && request.Format == ExcelFormat.Xlsx;

    /// <summary>
    /// 判断请求是否只要求原样保留模板。
    /// </summary>
    /// <param name="request">工作簿导出请求。</param>
    /// <returns>是否可以直接复制已经过预检的模板。</returns>
    internal static bool IsPreserveOnly(ExcelWorkbookExportRequest request) =>
        IsPreserveTemplate(request)
        && request.Sheets.Count == 0
        && request.LookupLists.Count == 0
        && request.CascadeLists.Count == 0
        && request.PivotTables.Count == 0
        && request.PivotTableUpdates.Count == 0
        && !request.MetadataSpecified;

    /// <summary>
    /// 同步复制模板快照。
    /// </summary>
    /// <param name="request">仅保留模板的工作簿导出请求。</param>
    /// <param name="destination">目标流。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="token">取消令牌。</param>
    internal static void Copy(ExcelWorkbookExportRequest request, Stream destination,
        string provider, CancellationToken token)
    {
        using var snapshot = ExcelPivotTemplateSnapshot.Create(request,
            request.ResourceLimits ?? new ExcelResourceLimits(), provider, token);
        var buffer = new byte[81920];
        int count;
        while ((count = snapshot.Source.Read(buffer, 0, buffer.Length)) > 0)
        {
            token.ThrowIfCancellationRequested();
            destination.Write(buffer, 0, count);
        }
        token.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 异步复制模板快照。
    /// </summary>
    /// <param name="request">仅保留模板的工作簿导出请求。</param>
    /// <param name="destination">目标流。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>表示复制操作的异步任务。</returns>
    internal static async Task CopyAsync(ExcelWorkbookExportRequest request, Stream destination,
        string provider, CancellationToken token)
    {
        using var snapshot = await ExcelPivotTemplateSnapshot.CreateAsync(request,
            request.ResourceLimits ?? new ExcelResourceLimits(), provider, token).ConfigureAwait(false);
        var buffer = new byte[81920];
        int count;
        while ((count = await snapshot.Source.ReadAsync(buffer.AsMemory(), token).ConfigureAwait(false)) > 0)
        {
            await destination.WriteAsync(buffer.AsMemory(0, count), token).ConfigureAwait(false);
        }
        await destination.FlushAsync(token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
    }
}
