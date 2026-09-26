using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.ProviderContract.Tests.Runner;
using Bing.Offices.Testing.Requests;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证所有适用 Provider 在预取消和真实异步 IO 中途均尊重调用方取消。
/// </summary>
public sealed class CancellationContractTest
{
    /// <summary>
    /// 预取消不得产生输出或导入结果，并应保留取消异常。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task PreCancelledAsync_ShouldStopBeforeIo(string provider)
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        var execution = await ProviderContractRunner.RoundTripAsync(
            ProviderDrivers.Get(provider),
            "Cancellation",
            ContractRequests.ScalarExport(),
            ContractRequests.ScalarImport(),
            ContractExecutionMode.Async,
            cancellation.Token);

        Assert.True(ContractActualOutcome.Exception == execution.ActualOutcome,
            execution.FormatAssertionFailure("cancellation outcome mismatch"));
        Assert.True(execution.Exception is OperationCanceledException,
            execution.FormatAssertionFailure("cancellation exception mismatch"));
        Assert.True(execution.Result == null, execution.FormatAssertionFailure("unexpected import result"));
        Assert.True(execution.Output.Length == 0, execution.FormatAssertionFailure("output is not empty"));
        Assert.True("OperationCanceled" == execution.ExpectedProfile,
            execution.FormatAssertionFailure("cancellation profile mismatch"));
        Assert.True("Xlsx" == execution.Format, execution.FormatAssertionFailure("format mismatch"));
    }

    /// <summary>
    /// 真实输入流首次读取后取消必须阻止三个 Provider 继续导入。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task MidReadCancellation_ShouldStopAfterRealIo(string provider)
    {
        using var generated = new MemoryStream();
        await ProviderDrivers.Get(provider).CreateExporter().ExportAsync(
            ContractRequests.ScalarExport(), generated);
        generated.Position = 0;
        using var cancellation = new CancellationTokenSource();
        using var source = new CancelAfterFirstReadStream(generated, cancellation);

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source,
                ContractRequests.ScalarImport(), cancellation.Token));

        Assert.True(source.HasRead);
        Assert.True(cancellation.IsCancellationRequested);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 获取参与契约测试的 Provider 名称。
    /// </summary>
    public static IEnumerable<object[]> Providers => ProviderDrivers.All.Select(driver => new object[] { driver.Name });

    /// <summary>
    /// 在第一次真实读取返回后取消调用方令牌的流包装器。
    /// </summary>
    private sealed class CancelAfterFirstReadStream : Stream
    {
        /// <summary>
        /// 提供测试输入的底层流，其所有权保留给调用方。
        /// </summary>
        private readonly Stream _inner;
        /// <summary>
        /// 首次读取到数据后需要触发的取消令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;
        /// <summary>
        /// 初始化一个 <see cref="CancelAfterFirstReadStream"/> 类型的实例。
        /// </summary>
        /// <param name="inner">提供输入内容的底层流。</param>
        /// <param name="cancellation">首次读取到数据后需要触发的取消令牌源。</param>
        internal CancelAfterFirstReadStream(Stream inner, CancellationTokenSource cancellation)
        {
            _inner = inner;
            _cancellation = cancellation;
        }
        /// <summary>
        /// 获取是否已读取到输入数据。
        /// </summary>
        internal bool HasRead { get; private set; }
        /// <inheritdoc />
        public override bool CanRead => _inner.CanRead;
        /// <inheritdoc />
        public override bool CanSeek => _inner.CanSeek;
        /// <inheritdoc />
        public override bool CanWrite => false;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        /// <inheritdoc />
        public override void Flush() => _inner.Flush();
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = _inner.Read(buffer, offset, count);
            Cancel(read);
            return read;
        }
        /// <inheritdoc />
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            var read = await _inner.ReadAsync(buffer, offset, count, cancellationToken);
            Cancel(read);
            return read;
        }
        /// <summary>
        /// 首次读取到数据后触发取消。
        /// </summary>
        /// <param name="read">本次读取的字节数。</param>
        private void Cancel(int read)
        {
            if (read > 0 && !HasRead)
            {
                HasRead = true;
                _cancellation.Cancel();
            }
        }
    }
}
