namespace Bing.Offices.Exceptions;

/// <summary>
/// 接收 Bing.Offices 公共运行异常的观察器。
/// </summary>
public interface IBingOfficesExceptionObserver
{
    /// <summary>
    /// 观察一个已经完成分类的公共异常。
    /// </summary>
    /// <param name="exception">公共异常。</param>
    void Observe(BingOfficesException exception);
}
