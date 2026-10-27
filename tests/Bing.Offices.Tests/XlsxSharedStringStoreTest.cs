using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Streaming;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>验证 XLSX 共享字符串存储的内存和磁盘边界。</summary>
public sealed class XlsxSharedStringStoreTest
{
    [Fact]
    public void Add_ShouldKeepExactBudgetInMemoryAndSpillWhenNextValueExceedsIt()
    {
        WithSession(4096, (session, _) =>
        {
            using var store = new XlsxSharedStringStore(session, 68, 64);

            store.Add("ab");

            Assert.Equal(1, store.Count);
            Assert.False(store.IsSpilled);
            Assert.Equal(68L, store.ResidentBytes);
            Assert.Equal("ab", store.Get(0));

            store.Add("x");

            Assert.Equal(2, store.Count);
            Assert.True(store.IsSpilled);
            Assert.Equal(0L, store.ResidentBytes);
            Assert.Equal("ab", store.Get(0));
            Assert.Equal("x", store.Get(1));
        });
    }

    [Fact]
    public void SpilledStore_ShouldSupportUtf8RandomReadsAndBoundLruCache()
    {
        WithSession(64 * 1024, (session, _) =>
        {
            using var store = new XlsxSharedStringStore(session, 140, 128);
            var values = new[] { "中", "😀", "", "café", new string('z', 80) };
            foreach (var value in values)
                store.Add(value);

            Assert.True(store.IsSpilled);
            Assert.Equal(values.Length, store.Count);
            Assert.Equal(values[3], store.Get(3));
            Assert.Equal(72L, store.ResidentBytes);
            Assert.Equal(values[0], store.Get(0));
            Assert.Equal(138L, store.ResidentBytes);
            var residentBytes = store.ResidentBytes;
            Assert.Equal(values[4], store.Get(4));
            Assert.Equal(residentBytes, store.ResidentBytes);
            Assert.Equal(values[1], store.Get(1));
            Assert.Equal(134L, store.ResidentBytes);
            Assert.Equal(values[0], store.Get(0));
            Assert.Equal(values[3], store.Get(3));
            Assert.Equal(138L, store.ResidentBytes);
            Assert.Equal(values[2], store.Get(2));
            Assert.Equal(136L, store.ResidentBytes);
        });
    }

    [Fact]
    public void SpillIndex_ShouldUseFixedTwelveByteRecordsWithoutResidentOffsetTable()
    {
        const int count = 100;
        WithSession(count * 13L, (session, _) =>
        {
            using var store = new XlsxSharedStringStore(session, 0, 8);
            foreach (var index in Enumerable.Range(0, count))
                store.Add("x");

            Assert.True(store.IsSpilled);
            Assert.Equal(count, store.Count);
            Assert.Equal(0L, store.ResidentBytes);
            Assert.Equal(count * 13L, session.PeakBytes);
            Assert.Equal("x", store.Get(count - 1));
            Assert.Equal(0L, store.ResidentBytes);
        });
    }

    [Fact]
    public void Add_ShouldRejectOverlongValueBeforeChangingStore()
    {
        WithSession(4096, (session, _) =>
        {
            using var store = new XlsxSharedStringStore(session, 4096, 2);

            Assert.Throws<InvalidDataException>(() => store.Add("中😀"));

            Assert.Equal(0, store.Count);
            Assert.False(store.IsSpilled);
            Assert.Equal(0L, store.ResidentBytes);
        });
    }

    [Fact]
    public void Spill_ShouldHonorSessionTemporaryFileBudgetAndKeepExistingMemoryValues()
    {
        WithSession(13, (session, _) =>
        {
            using var store = new XlsxSharedStringStore(session, 66, 16);
            store.Add("a");

            Assert.Throws<StreamingResourceLimitException>(() => store.Add("b"));

            Assert.Equal(1, store.Count);
            Assert.False(store.IsSpilled);
            Assert.Equal("a", store.Get(0));
            Assert.Equal(66L, store.ResidentBytes);
        });
    }

    [Fact]
    public void Append_ShouldRollbackDiskFilesWhenSessionBudgetIsExceeded()
    {
        WithSession(13, (session, _) =>
        {
            using var store = new XlsxSharedStringStore(session, 0, 16);
            store.Add("a");

            Assert.True(store.IsSpilled);
            Assert.Throws<StreamingResourceLimitException>(() => store.Add("b"));

            Assert.Equal(1, store.Count);
            Assert.Equal("a", store.Get(0));
            Assert.Throws<InvalidDataException>(() => store.Get(1));
        });
    }

    [Fact]
    public async Task AddAsync_ShouldUseDiskIoAndRespectCancellation()
    {
        await WithSessionAsync(4096, async (session, _) =>
        {
            using var store = new XlsxSharedStringStore(session, 0, 32);
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => store.AddAsync("cancel", cancelled.Token));
            Assert.Equal(0, store.Count);
            Assert.False(store.IsSpilled);

            await store.AddAsync("首", CancellationToken.None);
            await store.AddAsync("😀", CancellationToken.None);

            Assert.True(store.IsSpilled);
            Assert.Equal(2, store.Count);
            Assert.Equal("首", await store.GetAsync(0, CancellationToken.None));
            Assert.Equal("😀", await store.GetAsync(1, CancellationToken.None));
        });
    }

    [Fact]
    public async Task GetAsync_ShouldHonorCancellationAndRejectInvalidIndices()
    {
        await WithSessionAsync(4096, async (session, _) =>
        {
            using var store = new XlsxSharedStringStore(session, 0, 16);
            await store.AddAsync("value", CancellationToken.None);

            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => store.GetAsync(0, cancelled.Token));
            await Assert.ThrowsAsync<InvalidDataException>(
                () => store.GetAsync(1, CancellationToken.None));
            Assert.Throws<InvalidDataException>(() => store.Get(-1));
        });
    }

    private static void WithSession(long sessionBudget, Action<ExcelStreamingTempSession, string> action)
    {
        var root = Path.Combine(Path.GetTempPath(), "Bing.Offices.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            using (var session = new ExcelStreamingTempSession(root, sessionBudget))
                action(session, root);
            Assert.Empty(Directory.GetFileSystemEntries(root));
        }
        finally
        {
            if (Directory.Exists(root))
                Directory.Delete(root, true);
        }
    }

    private static async Task WithSessionAsync(long sessionBudget,
        Func<ExcelStreamingTempSession, string, Task> action)
    {
        var root = Path.Combine(Path.GetTempPath(), "Bing.Offices.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            using (var session = new ExcelStreamingTempSession(root, sessionBudget))
                await action(session, root);
            Assert.Empty(Directory.GetFileSystemEntries(root));
        }
        finally
        {
            if (Directory.Exists(root))
                Directory.Delete(root, true);
        }
    }
}
