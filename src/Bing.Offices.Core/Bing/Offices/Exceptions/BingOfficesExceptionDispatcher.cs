using System.Diagnostics;
using System.ComponentModel;

namespace Bing.Offices.Exceptions;

/// <summary>公共 Bing.Offices 异常观察分发器。</summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class BingOfficesExceptionDispatcher
{
    /// <summary>标记异常已完成观察，防止同一实例重复通知。</summary>
    public const string ObservedKey = "Bing.Offices.ExceptionObserved";
    /// <summary>保存观察器失败诊断的 Data 键。</summary>
    public const string ObserverFailureKey = "Bing.Offices.ExceptionObserverFailure";
    private readonly IReadOnlyList<IBingOfficesExceptionObserver> _observers;

    /// <summary>使用观察器集合创建异常分发器。</summary>
    /// <param name="observers">接收公共异常的观察器集合。</param>
    public BingOfficesExceptionDispatcher(IEnumerable<IBingOfficesExceptionObserver> observers = null)
    {
        _observers = new List<IBingOfficesExceptionObserver>(observers ?? Array.Empty<IBingOfficesExceptionObserver>());
    }

    /// <summary>向每个观察器发送异常，并保证同一异常实例最多通知一次。</summary>
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
