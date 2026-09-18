using System.Diagnostics;
using System.ComponentModel;

namespace Bing.Offices.Exceptions;

/// <summary>
/// 公共 Bing.Offices 异常观察分发器。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class BingOfficesExceptionDispatcher
{
    /// <summary>
    /// 写入异常 Data 的观察完成标记键，防止同一实例重复通知。
    /// </summary>
    public const string ObservedKey = "Bing.Offices.ExceptionObserved";
    /// <summary>
    /// 保存观察器失败诊断的异常 Data 键。
    /// </summary>
    public const string ObserverFailureKey = "Bing.Offices.ExceptionObserverFailure";
    /// <summary>
    /// 用于接收公共异常通知的观察器集合。
    /// </summary>
    private readonly IReadOnlyList<IBingOfficesExceptionObserver> _observers;

    /// <summary>
    /// 初始化一个 <see cref="BingOfficesExceptionDispatcher" /> 类型的实例。
    /// </summary>
    /// <param name="observers">接收公共异常的观察器集合。</param>
    public BingOfficesExceptionDispatcher(IEnumerable<IBingOfficesExceptionObserver> observers = null)
    {
        _observers = new List<IBingOfficesExceptionObserver>(observers ?? Array.Empty<IBingOfficesExceptionObserver>());
    }

    /// <summary>
    /// 向每个观察器发送异常，并保证同一异常实例最多通知一次。
    /// </summary>
    /// <param name="exception">待观察的公共异常。</param>
    public void Observe(BingOfficesException exception)
    {
        if (exception == null)
            return;
        lock (exception)
        {
            if (exception.Data.Contains(ObservedKey))
                return;
            exception.Data[ObservedKey] = true;
        }
        foreach (var observer in _observers)
        {
            if (observer == null)
                continue;
            try
            {
                observer.Observe(exception);
            }
            catch (Exception observerException) when (observerException is not OutOfMemoryException
                && observerException is not StackOverflowException)
            {
                AddObserverFailure(exception, observerException);
            }
        }
    }

    /// <summary>
    /// 记录异常观察器失败，避免观察器异常覆盖原始异常。
    /// </summary>
    /// <param name="exception">正在通知观察器的公共异常。</param>
    /// <param name="observerException">观察器抛出的异常。</param>
    private static void AddObserverFailure(BingOfficesException exception, Exception observerException)
    {
        lock (exception)
        {
            if (!exception.Data.Contains(ObserverFailureKey))
            {
                exception.Data[ObserverFailureKey] = observerException;
                return;
            }
        }
        Trace.WriteLine($"Bing.Offices 异常观察器失败: {observerException.GetType().Name}");
    }
}
