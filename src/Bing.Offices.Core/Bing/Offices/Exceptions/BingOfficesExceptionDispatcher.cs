using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.ComponentModel;

namespace Bing.Offices.Exceptions;

/// <summary>公共 Bing.Offices 异常观察分发器。</summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class BingOfficesExceptionDispatcher
{
    public const string ObservedKey = "Bing.Offices.ExceptionObserved";
    public const string ObserverFailureKey = "Bing.Offices.ExceptionObserverFailure";
    private readonly IReadOnlyList<IBingOfficesExceptionObserver> _observers;

    public BingOfficesExceptionDispatcher(IEnumerable<IBingOfficesExceptionObserver> observers = null)
    {
        _observers = new List<IBingOfficesExceptionObserver>(observers ?? Array.Empty<IBingOfficesExceptionObserver>());
    }

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
