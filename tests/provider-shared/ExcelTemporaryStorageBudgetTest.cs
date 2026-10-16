using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Entities;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 实体导入临时存储共享预算测试。
/// </summary>
public sealed class ExcelTemporaryStorageBudgetTest
{
    /// <summary>
    /// 验证预算在多个同时存活文件间聚合，并允许恰好达到上限。
    /// </summary>
    [Fact]
    public void CreateFile_WhenAggregateReachesLimit_ShouldRejectOnlyFurtherGrowth()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            var budget = new ExcelTemporaryStorageBudget(5, "TestProvider");
            using var first = CreateFile(budget, directory, "first.bin");
            using var second = CreateFile(budget, directory, "second.bin");

            first.Write(new byte[3], 0, 3);
            second.Write(new byte[2].AsSpan());

            var exception = Assert.Throws<BingOfficesResourceLimitException>(
                () => second.WriteByte(0x7f));
            Assert.Equal("TestProvider", exception.Provider);
            Assert.Equal(BingOfficesOperation.Import, exception.Operation);
            Assert.Equal(BingOfficesStage.Open, exception.Stage);
            Assert.Equal(2, second.Length);

            second.Dispose();
            first.Position = first.Length;
            first.WriteByte(0x7f);
            Assert.Equal(4, first.Length);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证覆盖写入不重复计数，定位扩容和 SetLength 遵循当前逻辑长度。
    /// </summary>
    [Fact]
    public void WriteAndSetLength_ShouldAccountForGrowthOverwriteShrinkAndRelease()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            var budget = new ExcelTemporaryStorageBudget(8, "TestProvider");
            using var first = CreateFile(budget, directory, "first.bin");
            first.Write(new byte[6], 0, 6);

            first.Position = 2;
            first.Write(new byte[] { 1, 2 }.AsSpan());
            first.Position = 6;
            first.WriteByte(3);
            Assert.Equal(7, first.Length);

            Assert.Throws<BingOfficesResourceLimitException>(() => first.SetLength(9));
            Assert.Equal(7, first.Length);
            first.SetLength(4);

            using var second = CreateFile(budget, directory, "second.bin");
            second.Write(new byte[4], 0, 4);
            first.Position = 4;
            Assert.Throws<BingOfficesResourceLimitException>(() => first.WriteByte(4));
            Assert.Equal(4, first.Length);

            second.SetLength(1);
            first.WriteByte(4);
            Assert.Equal(5, first.Length);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证真实异步写入计入预算，预取消不会预留字节，异步释放可供其他文件使用。
    /// </summary>
    [Fact]
    public async Task WriteAsync_ShouldAccountCancelAndReleaseAcrossFiles()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            var budget = new ExcelTemporaryStorageBudget(6, "TestProvider");
            await using var first = CreateFile(budget, directory, "first.bin");
            await using var second = CreateFile(budget, directory, "second.bin");

            await first.WriteAsync(new byte[4].AsMemory());
            await Assert.ThrowsAsync<BingOfficesResourceLimitException>(async () =>
                await second.WriteAsync(new byte[3].AsMemory()));

            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
                await second.WriteAsync(new byte[1].AsMemory(), cancellation.Token));

            await first.DisposeAsync();
            await second.WriteAsync(new byte[6].AsMemory());
            Assert.Equal(6, second.Length);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证定位到预算范围外后的零字节写入不会预留空间或扩展文件。
    /// </summary>
    [Fact]
    public async Task ZeroLengthWrite_WhenPositionExceedsLimit_ShouldNotReserveOrGrow()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            var budget = new ExcelTemporaryStorageBudget(4, "TestProvider");
            using var first = CreateFile(budget, directory, "empty.bin");
            WriteZeroLengthSync(first);
            await first.WriteAsync(ReadOnlyMemory<byte>.Empty);
            Assert.Equal(0, first.Length);

            using var second = CreateFile(budget, directory, "full.bin");
            second.Write(new byte[4], 0, 4);
            Assert.Equal(4, second.Length);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证并发文件共享同一预算，而不同导入预算彼此隔离。
    /// </summary>
    [Fact]
    public async Task CreateFile_WhenUsedConcurrently_ShouldShareOnlyItsOwnBudget()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            var sharedBudget = new ExcelTemporaryStorageBudget(4, "SharedProvider");
            using var first = CreateFile(sharedBudget, directory, "shared-first.bin");
            using var second = CreateFile(sharedBudget, directory, "shared-second.bin");
            var sharedResults = await Task.WhenAll(
                Task.Run(() => TryWriteFourBytes(first)),
                Task.Run(() => TryWriteFourBytes(second)));
            Assert.Single(sharedResults.Where(result => result));

            var firstBudget = new ExcelTemporaryStorageBudget(4, "FirstProvider");
            var secondBudget = new ExcelTemporaryStorageBudget(4, "SecondProvider");
            using var isolatedFirst = CreateFile(firstBudget, directory, "isolated-first.bin");
            using var isolatedSecond = CreateFile(secondBudget, directory, "isolated-second.bin");
            var isolatedResults = await Task.WhenAll(
                Task.Run(() => TryWriteFourBytes(isolatedFirst)),
                Task.Run(() => TryWriteFourBytes(isolatedSecond)));
            Assert.All(isolatedResults, result => Assert.True(result));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证有限预算必须为正数。
    /// </summary>
    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Constructor_WhenMaximumIsNotPositive_ShouldThrow(long maximumBytes)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelTemporaryStorageBudget(maximumBytes, "TestProvider"));
    }

    private static Stream CreateFile(ExcelTemporaryStorageBudget budget, string directory,
        string fileName) => budget.CreateFile(Path.Combine(directory, fileName),
        FileOptions.Asynchronous | FileOptions.DeleteOnClose);

    private static bool TryWriteFourBytes(Stream stream)
    {
        try
        {
            stream.Write(new byte[4], 0, 4);
            return true;
        }
        catch (BingOfficesResourceLimitException)
        {
            return false;
        }
    }

    private static void WriteZeroLengthSync(Stream stream)
    {
        stream.Position = 16;
        stream.Write(Array.Empty<byte>(), 0, 0);
        stream.Write(ReadOnlySpan<byte>.Empty);
    }

    private static string CreateTemporaryDirectory()
    {
        var directory = Path.Combine(Path.GetTempPath(),
            $"bing-offices-entity-temp-budget-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        return directory;
    }
}
