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

/// <summary>
/// 以并发和排队上限限制 ClosedXML Workbook DOM 的创建。
/// </summary>
internal sealed class ClosedXmlWorkbookAdmission : IClosedXmlWorkbookAdmission
{
    /// <summary>
    /// 直接构造 Provider 时使用的默认准入器。
    /// </summary>
    private static readonly ClosedXmlWorkbookAdmission Default =
        new(new ClosedXmlProviderOptions());

    /// <summary>
    /// 控制实际 Workbook 并发数的信号量。
    /// </summary>
    private readonly SemaphoreSlim _gate;

    /// <summary>
    /// 调用方请求的最大并发 Workbook 数量。
    /// </summary>
    private readonly int _requestedConcurrency;

    /// <summary>
    /// 允许等待准入的最大排队数量。
    /// </summary>
    private readonly int _maxQueued;

    /// <summary>
    /// 当前等待准入的操作数量。
    /// </summary>
    private int _queued;

    /// <summary>
    /// 当前持有准入令牌的操作数量。
    /// </summary>
    private int _active;

    /// <summary>
    /// 观测到的最大实际并行 Workbook 数量。
    /// </summary>
    private int _maxActualParallelism;

    /// <summary>
    /// 观测到的最大排队操作数量。
    /// </summary>
    private int _maxObservedQueued;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlWorkbookAdmission" /> 类型的实例。
    /// </summary>
    /// <param name="options">并发和排队限制。</param>
    internal ClosedXmlWorkbookAdmission(ClosedXmlProviderOptions options)
    {
        options ??= new ClosedXmlProviderOptions();
        options.Validate();
        _requestedConcurrency = options.MaxConcurrentWorkbooks;
        _gate = new SemaphoreSlim(options.MaxConcurrentWorkbooks, options.MaxConcurrentWorkbooks);
        _maxQueued = options.MaxQueuedOperations;
    }

    /// <summary>
    /// 获取直接构造 Provider 时使用的进程级默认准入器。
    /// </summary>
    internal static IClosedXmlWorkbookAdmission SharedDefault => Default;

    /// <summary>
    /// 获取调用方配置的最大并发数。
    /// </summary>
    internal int RequestedConcurrency => _requestedConcurrency;

    /// <summary>
    /// 获取允许的最大排队数量。
    /// </summary>
    internal int MaxQueuedOperations => _maxQueued;

    /// <summary>
    /// 获取观测到的最大实际并行数。
    /// </summary>
    internal int MaxActualParallelism => Volatile.Read(ref _maxActualParallelism);

    /// <summary>
    /// 获取观测到的最大排队数量。
    /// </summary>
    internal int MaxObservedQueuedOperations => Volatile.Read(ref _maxObservedQueued);

    /// <inheritdoc />
    public IDisposable Acquire(CancellationToken cancellationToken, BingOfficesOperation operation)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (_gate.Wait(0))
            return CreateReleaser();
        EnterQueue(operation);
        try
        {
            _gate.Wait(cancellationToken);
            return CreateReleaser();
        }
        finally
        {
            Interlocked.Decrement(ref _queued);
        }
    }

    /// <inheritdoc />
    public async ValueTask<IAsyncDisposable> AcquireAsync(CancellationToken cancellationToken,
        BingOfficesOperation operation)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (_gate.Wait(0))
            return CreateAsyncReleaser();
        EnterQueue(operation);
        try
        {
            await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            return CreateAsyncReleaser();
        }
        finally
        {
            Interlocked.Decrement(ref _queued);
        }
    }

    /// <summary>
    /// 将当前操作加入排队并检查排队上限。
    /// </summary>
    /// <param name="operation">触发准入的操作类型。</param>
    private void EnterQueue(BingOfficesOperation operation)
    {
        var queued = Interlocked.Increment(ref _queued);
        UpdateMaximum(ref _maxObservedQueued, queued);
        if (queued <= _maxQueued)
            return;
        Interlocked.Decrement(ref _queued);
        throw new BingOfficesResourceLimitException(
            $"ClosedXML Workbook 准入队列超过限制: {_maxQueued}。",
            provider: "ClosedXML", operation: operation, stage: BingOfficesStage.Preflight);
    }

    /// <summary>
    /// 创建同步准入令牌并记录活动并发数。
    /// </summary>
    /// <returns>同步释放令牌。</returns>
    private Releaser CreateReleaser()
    {
        EnterActive();
        return new Releaser(this, _gate);
    }

    /// <summary>
    /// 创建异步准入令牌并记录活动并发数。
    /// </summary>
    /// <returns>异步释放令牌。</returns>
    private AsyncReleaser CreateAsyncReleaser()
    {
        EnterActive();
        return new AsyncReleaser(this, _gate);
    }

    /// <summary>
    /// 增加当前活动 Workbook 数量并更新峰值。
    /// </summary>
    private void EnterActive()
    {
        var active = Interlocked.Increment(ref _active);
        UpdateMaximum(ref _maxActualParallelism, active);
    }

    /// <summary>
    /// 减少当前活动 Workbook 数量。
    /// </summary>
    private void ExitActive() => Interlocked.Decrement(ref _active);

    /// <summary>
    /// 以无锁方式更新一个整数峰值。
    /// </summary>
    /// <param name="target">待更新的峰值字段。</param>
    /// <param name="value">候选值。</param>
    private static void UpdateMaximum(ref int target, int value)
    {
        while (true)
        {
            var current = Volatile.Read(ref target);
            if (current >= value || Interlocked.CompareExchange(ref target, value, current) == current)
                return;
        }
    }

    /// <summary>
    /// 释放同步准入额度的令牌。
    /// </summary>
    private sealed class Releaser : IDisposable
    {
        /// <summary>
        /// 持有该令牌的准入器。
        /// </summary>
        private readonly ClosedXmlWorkbookAdmission _owner;

        /// <summary>
        /// 准入信号量。
        /// </summary>
        private readonly SemaphoreSlim _gate;

        /// <summary>
        /// 防止令牌被重复释放。
        /// </summary>
        private int _released;

        /// <summary>
        /// 初始化一个 <see cref="Releaser" /> 类型的实例。
        /// </summary>
        /// <param name="owner">准入器。</param>
        /// <param name="gate">准入信号量。</param>
        public Releaser(ClosedXmlWorkbookAdmission owner, SemaphoreSlim gate)
        {
            _owner = owner;
            _gate = gate;
        }

        /// <inheritdoc />
        public void Dispose()
        {
            if (Interlocked.Exchange(ref _released, 1) == 0)
            {
                _owner.ExitActive();
                _gate.Release();
            }
        }
    }

    /// <summary>
    /// 释放异步准入额度的令牌。
    /// </summary>
    private sealed class AsyncReleaser : IAsyncDisposable
    {
        /// <summary>
        /// 持有该令牌的准入器。
        /// </summary>
        private readonly ClosedXmlWorkbookAdmission _owner;

        /// <summary>
        /// 准入信号量。
        /// </summary>
        private readonly SemaphoreSlim _gate;

        /// <summary>
        /// 防止令牌被重复释放。
        /// </summary>
        private int _released;

        /// <summary>
        /// 初始化一个 <see cref="AsyncReleaser" /> 类型的实例。
        /// </summary>
        /// <param name="owner">准入器。</param>
        /// <param name="gate">准入信号量。</param>
        public AsyncReleaser(ClosedXmlWorkbookAdmission owner, SemaphoreSlim gate)
        {
            _owner = owner;
            _gate = gate;
        }

        /// <inheritdoc />
        public ValueTask DisposeAsync()
        {
            if (Interlocked.Exchange(ref _released, 1) == 0)
            {
                _owner.ExitActive();
                _gate.Release();
            }
            return ValueTask.CompletedTask;
        }
    }
}
