using System.Threading;
using Bing.Offices.Exceptions;

namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// 控制 ClosedXML Workbook DOM 创建并提供同步、异步准入入口。
/// </summary>
internal interface IClosedXmlWorkbookAdmission
{
    /// <summary>
    /// 同步获取一个 Workbook 准入令牌。
    /// </summary>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="operation">触发准入的操作类型。</param>
    /// <returns>释放后归还准入额度的令牌。</returns>
    IDisposable Acquire(CancellationToken cancellationToken, BingOfficesOperation operation);

    /// <summary>
    /// 异步获取一个 Workbook 准入令牌。
    /// </summary>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="operation">触发准入的操作类型。</param>
    /// <returns>异步准入令牌。</returns>
    ValueTask<IAsyncDisposable> AcquireAsync(CancellationToken cancellationToken,
        BingOfficesOperation operation);
}
