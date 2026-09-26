using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.Testing.Models;
using Bing.Offices.Testing.Requests;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证失败工作簿在真实文件输入和调用方流边界上的提交语义。
/// </summary>
public sealed class FailureWorkbookFileContractTest
{
    /// <summary>
    /// 启用失败工作簿时，流目标与路径目标必须严格互斥。
    /// </summary>
    [Fact]
    public void FailureWorkbookTargets_ShouldRequireExactlyOneDestination()
    {
        using var destination = new MemoryStream();

        Assert.Throws<ArgumentException>(() => new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal
        }.Validate());
        Assert.Throws<ArgumentException>(() => new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            Destination = destination,
            DestinationPath = "failure.xlsx"
        }.Validate());
        Assert.Throws<ArgumentException>(() => new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            DestinationPath = "   "
        }.Validate());
    }

    /// <summary>
    /// 真实文件输入应生成失败工作簿，或按档案在预检阶段拒绝。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task AnnotatedOriginal_FromRealFile_ShouldCommitOnlyCompleteOutput(string provider)
    {
        var path = CreateTemporaryPath();
        try
        {
            await ProviderDrivers.Get(provider).CreateExporter().ExportToFileAsync(
                ContractRequests.ValidationExport(), path);
            using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
            using var destination = new MemoryStream();
            var request = CreateRequest(destination);

            if (provider == "MiniExcel")
            {
                var exception = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
                    ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request));
                Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
                Assert.Equal(0, destination.Length);
            }
            else
            {
                var result = await ProviderDrivers.Get(provider).CreateImporter()
                    .ImportAsync(source, request);
                Assert.False(result.IsSuccess);
                Assert.NotEmpty(result.Errors);
                Assert.True(destination.Length > 0);
            }

            Assert.True(source.CanRead);
        }
        finally
        {
            DeleteTemporaryPath(path);
        }
    }

    /// <summary>
    /// 真实文件输入被取消时，失败输出目标应保持原有内容且输入流不被关闭。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task AnnotatedOriginal_CancellationFromRealFile_ShouldPreserveDestination(string provider)
    {
        var path = CreateTemporaryPath();
        var sentinel = Encoding.UTF8.GetBytes("existing-failure-output");
        try
        {
            await ProviderDrivers.Get(provider).CreateExporter().ExportToFileAsync(
                ContractRequests.ValidationExport(), path);
            using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
            using var destination = new MemoryStream();
            await destination.WriteAsync(sentinel, 0, sentinel.Length);
            destination.Position = 0;
            var request = CreateRequest(destination);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request,
                    cancellation.Token));

            Assert.Equal(sentinel, destination.ToArray());
            Assert.True(source.CanRead);
        }
        finally
        {
            DeleteTemporaryPath(path);
        }
    }

    /// <summary>
    /// 路径目标应在支持的 Provider 中原子替换，在 MiniExcel 中预检拒绝且不创建文件。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task AnnotatedOriginal_PathDestination_ShouldReplaceOnlyCompleteOutput(string provider)
    {
        var sourcePath = CreateTemporaryPath();
        var destinationPath = CreateTemporaryPath();
        var sentinel = Encoding.UTF8.GetBytes("existing-failure-output");
        try
        {
            await ProviderDrivers.Get(provider).CreateExporter().ExportToFileAsync(
                ContractRequests.ValidationExport(), sourcePath);
            await File.WriteAllBytesAsync(destinationPath, sentinel);
            await using var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read,
                4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
            var request = ExcelImport.Workbook<ValidationContractWorkbook>(builder => builder
                .FailureWorkbook(new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                    DestinationPath = destinationPath
                })
                .Sheet<ValidationContractRow>("Data", root => root.Rows,
                    sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

            if (provider == "MiniExcel")
            {
                await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
                    ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request));
                Assert.Equal(sentinel, await File.ReadAllBytesAsync(destinationPath));
            }
            else
            {
                var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request);
                Assert.False(result.IsSuccess);
                Assert.NotEqual(sentinel, await File.ReadAllBytesAsync(destinationPath));
            }
        }
        finally
        {
            DeleteTemporaryPath(sourcePath);
            DeleteTemporaryPath(destinationPath);
        }
    }

    /// <summary>
    /// 同步路径目标应创建新文件并在提交后不遗留同目录临时文件。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task AnnotatedOriginal_PathDestination_SynchronouslyShouldCreateAtomicOutput(string provider)
    {
        var sourcePath = CreateTemporaryPath();
        var destinationPath = CreateTemporaryPath();
        try
        {
            await ProviderDrivers.Get(provider).CreateExporter().ExportToFileAsync(
                ContractRequests.ValidationExport(), sourcePath);
            using var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            var request = ExcelImport.Workbook<ValidationContractWorkbook>(builder => builder
                .FailureWorkbook(new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                    DestinationPath = destinationPath
                })
                .Sheet<ValidationContractRow>("Data", root => root.Rows,
                    sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

            var result = ProviderDrivers.Get(provider).CreateImporter().Import(source, request);

            Assert.False(result.IsSuccess);
            Assert.True(File.Exists(destinationPath));
            Assert.NotEmpty(await File.ReadAllBytesAsync(destinationPath));
            Assert.Empty(FindCommitTemporaryFiles(destinationPath));
        }
        finally
        {
            DeleteTemporaryPath(sourcePath);
            DeleteTemporaryPath(destinationPath);
        }
    }

    /// <summary>
    /// 没有导入错误时，失败工作簿路径目标不得创建或替换任何文件。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task SuccessfulImport_ShouldLeaveFailurePathUntouched(string provider)
    {
        var sourcePath = CreateTemporaryPath();
        var destinationPath = CreateTemporaryPath();
        var sentinel = Encoding.UTF8.GetBytes("existing-success-output");
        try
        {
            await ProviderDrivers.Get(provider).CreateExporter().ExportToFileAsync(
                ContractRequests.ScalarExport(), sourcePath);
            await File.WriteAllBytesAsync(destinationPath, sentinel);
            await using var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read,
                4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
            var request = ExcelImport.Workbook<ScalarContractWorkbook>(builder => builder
                .FailureWorkbook(new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                    DestinationPath = destinationPath
                })
                .Sheet<ScalarContractRow>("Data", root => root.Rows));

            var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request);

            Assert.True(result.IsSuccess);
            Assert.Equal(sentinel, await File.ReadAllBytesAsync(destinationPath));
            Assert.Empty(FindCommitTemporaryFiles(destinationPath));
        }
        finally
        {
            DeleteTemporaryPath(sourcePath);
            DeleteTemporaryPath(destinationPath);
        }
    }

    /// <summary>
    /// 输入已开始真实读取后取消时，路径目标必须保留旧文件且清理提交临时文件。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task PathDestination_MidReadCancellation_ShouldPreserveExistingFile(string provider)
    {
        var sourcePath = CreateTemporaryPath();
        var destinationPath = CreateTemporaryPath();
        var sentinel = Encoding.UTF8.GetBytes("existing-cancel-output");
        try
        {
            await ProviderDrivers.Get(provider).CreateExporter().ExportToFileAsync(
                ContractRequests.ValidationExport(), sourcePath);
            await File.WriteAllBytesAsync(destinationPath, sentinel);
            using var cancellation = new CancellationTokenSource();
            await using var file = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read,
                4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
            using var source = new CancelAfterFirstReadStream(file, cancellation);
            var request = ExcelImport.Workbook<ValidationContractWorkbook>(builder => builder
                .FailureWorkbook(new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                    DestinationPath = destinationPath
                })
                .Sheet<ValidationContractRow>("Data", root => root.Rows,
                    sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request, cancellation.Token));

            Assert.True(source.HasRead);
            Assert.Equal(sentinel, await File.ReadAllBytesAsync(destinationPath));
            Assert.Empty(FindCommitTemporaryFiles(destinationPath));
        }
        finally
        {
            DeleteTemporaryPath(sourcePath);
            DeleteTemporaryPath(destinationPath);
        }
    }

    /// <summary>
    /// 创建公共失败工作簿请求。
    /// </summary>
    /// <param name="destination">接收失败工作簿的目标流。</param>
    /// <returns>启用原始工作簿批注模式的导入请求。</returns>
    private static ExcelWorkbookImportRequest<ValidationContractWorkbook> CreateRequest(Stream destination) =>
        ExcelImport.Workbook<ValidationContractWorkbook>(builder => builder
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination
            })
            .Sheet<ValidationContractRow>("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

    /// <summary>
    /// 创建本测试专用的临时 XLSX 路径。
    /// </summary>
    /// <returns>当前测试独占的临时 XLSX 文件路径。</returns>
    private static string CreateTemporaryPath() => Path.Combine(Path.GetTempPath(),
        $"bing-offices-contract-{Guid.NewGuid():N}.xlsx");

    /// <summary>
    /// 删除本测试创建的临时文件。
    /// </summary>
    /// <param name="path">待删除的测试临时文件路径。</param>
    private static void DeleteTemporaryPath(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }

    /// <summary>
    /// 查找本次目标路径旁可能遗留的原子提交临时文件。
    /// </summary>
    /// <param name="destinationPath">失败工作簿目标文件路径。</param>
    /// <returns>目标路径所在目录中的提交临时文件路径集合。</returns>
    private static string[] FindCommitTemporaryFiles(string destinationPath) => Directory.GetFiles(
        Path.GetDirectoryName(destinationPath)!, Path.GetFileName(destinationPath) + ".*.tmp");

    /// <summary>
    /// 在首个同步或异步读取返回后取消令牌，确保取消发生于真实 IO 之后。
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
            CancelAfterRead(read);
            return read;
        }
        /// <inheritdoc />
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            var read = await _inner.ReadAsync(buffer, cancellationToken);
            CancelAfterRead(read);
            return read;
        }
        /// <summary>
        /// 首次读取到数据后触发取消。
        /// </summary>
        /// <param name="read">本次读取的字节数。</param>
        private void CancelAfterRead(int read)
        {
            if (read > 0 && !HasRead)
            {
                HasRead = true;
                _cancellation.Cancel();
            }
        }
    }

    /// <summary>
    /// 获取参与契约测试的 Provider 名称。
    /// </summary>
    public static System.Collections.Generic.IEnumerable<object[]> Providers =>
        ProviderDrivers.All.Select(driver => new object[] { driver.Name });
}
