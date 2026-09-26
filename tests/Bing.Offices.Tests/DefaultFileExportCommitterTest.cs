using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.IO;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 默认文件导出提交实现的职责级测试。
/// </summary>
public sealed class DefaultFileExportCommitterTest
{
    /// <summary>
    /// 测试 - 新目标应先写入同目录临时文件，再移动为最终文件。
    /// </summary>
    [Fact]
    public void Commit_NewTarget_ShouldWriteAndMove()
    {
        var path = CreatePath();
        try
        {
            new DefaultFileExportCommitter().Commit(path, stream => WriteText(stream, "新内容"),
                CancellationToken.None, "CSV");

            Assert.Equal("新内容", File.ReadAllText(path, Encoding.UTF8));
            Assert.Empty(GetTemporaryFiles(path));
        }
        finally
        {
            DeleteFile(path);
        }
    }

    /// <summary>
    /// 测试 - 已存在目标只有在新内容完成写入后才应被替换。
    /// </summary>
    [Fact]
    public void Commit_ExistingTarget_ShouldReplaceAfterSuccessfulWrite()
    {
        var path = CreatePath();
        File.WriteAllText(path, "旧内容", Encoding.UTF8);
        try
        {
            new DefaultFileExportCommitter().Commit(path, stream => WriteText(stream, "新内容"),
                CancellationToken.None, "Excel");

            Assert.Equal("新内容", File.ReadAllText(path, Encoding.UTF8));
            Assert.Empty(GetTemporaryFiles(path));
        }
        finally
        {
            DeleteFile(path);
        }
    }

    /// <summary>
    /// 测试 - 内容生成异常应原样传播，并保留既有目标且清理临时文件。
    /// </summary>
    [Fact]
    public void Commit_WriteFailure_ShouldPreserveExceptionAndTarget()
    {
        var path = CreatePath();
        File.WriteAllText(path, "旧内容", Encoding.UTF8);
        var expected = new InvalidOperationException("写入失败");
        try
        {
            var actual = Assert.Throws<InvalidOperationException>(() =>
                new DefaultFileExportCommitter().Commit(path, _ => throw expected,
                    CancellationToken.None, "CSV"));

            Assert.Same(expected, actual);
            Assert.Equal("旧内容", File.ReadAllText(path, Encoding.UTF8));
            Assert.Empty(GetTemporaryFiles(path));
        }
        finally
        {
            DeleteFile(path);
        }
    }

    /// <summary>
    /// 测试 - 预取消应在创建临时文件前终止，并且不产生目标或 staging 文件。
    /// </summary>
    [Fact]
    public void Commit_PreCanceled_ShouldNotCreateFiles()
    {
        var path = CreatePath();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        try
        {
            Assert.Throws<OperationCanceledException>(() =>
                new DefaultFileExportCommitter().Commit(path, stream => WriteText(stream, "不会写入"),
                    cancellation.Token, "CSV"));

            Assert.False(File.Exists(path));
            Assert.Empty(GetTemporaryFiles(path));
        }
        finally
        {
            DeleteFile(path);
        }
    }

    /// <summary>
    /// 验证写入后取消保留目标文件状态并清理临时文件。
    /// </summary>
    /// <param name="existingTarget">是否预先创建需要保留的目标文件。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Commit_CanceledAfterWrite_ShouldPreserveTargetAndCleanTemporaryFile(bool existingTarget)
    {
        var path = CreatePath();
        var sentinel = Encoding.UTF8.GetBytes("existing-target");
        if (existingTarget)
            File.WriteAllBytes(path, sentinel);
        using var cancellation = new CancellationTokenSource();
        try
        {
            Assert.Throws<OperationCanceledException>(() =>
                new DefaultFileExportCommitter().Commit(path, stream =>
                {
                    WriteText(stream, "部分内容");
                    cancellation.Cancel();
                }, cancellation.Token, "CSV"));

            if (existingTarget)
                Assert.Equal(sentinel, File.ReadAllBytes(path));
            else
                Assert.False(File.Exists(path));
            Assert.Empty(GetTemporaryFiles(path));
        }
        finally
        {
            DeleteFile(path);
        }
    }

    /// <summary>
    /// 验证异步提交新目标时会写入临时文件并移动为目标文件。
    /// </summary>
    [Fact]
    public async Task CommitAsync_NewTarget_ShouldWriteAndMove()
    {
        var path = CreatePath();
        try
        {
            await new DefaultFileExportCommitter().CommitAsync(path, async (stream, token) =>
            {
                var bytes = Encoding.UTF8.GetBytes("异步内容");
                await stream.WriteAsync(bytes, 0, bytes.Length, token);
            }, CancellationToken.None, "CSV");

            Assert.Equal("异步内容", File.ReadAllText(path, Encoding.UTF8));
            Assert.Empty(GetTemporaryFiles(path));
        }
        finally
        {
            DeleteFile(path);
        }
    }

    /// <summary>
    /// 验证异步提交已有目标时会在成功写入后替换旧文件。
    /// </summary>
    [Fact]
    public async Task CommitAsync_ExistingTarget_ShouldReplaceAfterSuccessfulWrite()
    {
        var path = CreatePath();
        File.WriteAllText(path, "旧内容", Encoding.UTF8);
        try
        {
            await new DefaultFileExportCommitter().CommitAsync(path, async (stream, token) =>
            {
                var bytes = Encoding.UTF8.GetBytes("新异步内容");
                await stream.WriteAsync(bytes, 0, bytes.Length, token);
            }, CancellationToken.None, "Excel");

            Assert.Equal("新异步内容", File.ReadAllText(path, Encoding.UTF8));
            Assert.Empty(GetTemporaryFiles(path));
        }
        finally
        {
            DeleteFile(path);
        }
    }

    /// <summary>
    /// 验证异步提交写入失败时会保留异常和已有目标。
    /// </summary>
    [Fact]
    public async Task CommitAsync_WriteFailure_ShouldPreserveExceptionAndTarget()
    {
        var path = CreatePath();
        File.WriteAllText(path, "旧内容", Encoding.UTF8);
        var expected = new InvalidOperationException("异步写入失败");
        try
        {
            var actual = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                new DefaultFileExportCommitter().CommitAsync(path, (stream, token) =>
                    Task.FromException(expected), CancellationToken.None, "CSV"));

            Assert.Same(expected, actual);
            Assert.Equal("旧内容", File.ReadAllText(path, Encoding.UTF8));
            Assert.Empty(GetTemporaryFiles(path));
        }
        finally
        {
            DeleteFile(path);
        }
    }

    /// <summary>
    /// 验证异步提交预先取消时不会创建文件。
    /// </summary>
    [Fact]
    public async Task CommitAsync_PreCanceled_ShouldNotCreateFiles()
    {
        var path = CreatePath();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        try
        {
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                new DefaultFileExportCommitter().CommitAsync(path,
                    (stream, token) => Task.CompletedTask, cancellation.Token, "CSV"));

            Assert.False(File.Exists(path));
            Assert.Empty(GetTemporaryFiles(path));
        }
        finally
        {
            DeleteFile(path);
        }
    }

    /// <summary>
    /// 验证写入后取消保留目标文件状态并清理临时文件。
    /// </summary>
    /// <param name="existingTarget">是否预先创建需要保留的目标文件。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CommitAsync_CanceledAfterWrite_ShouldCleanTemporaryFile(bool existingTarget)
    {
        var path = CreatePath();
        var sentinel = Encoding.UTF8.GetBytes("existing-target");
        if (existingTarget)
            File.WriteAllBytes(path, sentinel);
        using var cancellation = new CancellationTokenSource();
        try
        {
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                new DefaultFileExportCommitter().CommitAsync(path, async (stream, token) =>
                {
                    var bytes = Encoding.UTF8.GetBytes("部分内容");
                    await stream.WriteAsync(bytes, 0, bytes.Length, token);
                    cancellation.Cancel();
                }, cancellation.Token, "CSV"));

            if (existingTarget)
                Assert.Equal(sentinel, File.ReadAllBytes(path));
            else
                Assert.False(File.Exists(path));
            Assert.Empty(GetTemporaryFiles(path));
        }
        finally
        {
            DeleteFile(path);
        }
    }

    /// <summary>
    /// 创建测试路径。
    /// </summary>
    /// <returns>系统临时目录内的唯一测试文件路径。</returns>
    private static string CreatePath()
        => Path.Combine(Path.GetTempPath(), $"Bing.Offices.DefaultCommitter.{Guid.NewGuid():N}.tmp");

    /// <summary>
    /// 获取临时文件。
    /// </summary>
    /// <param name="path">目标文件或目录路径。</param>
    /// <returns>与目标文件对应的原子提交临时文件路径数组。</returns>
    private static string[] GetTemporaryFiles(string path)
        => Directory.GetFiles(Path.GetDirectoryName(path), Path.GetFileName(path) + ".*.tmp");

    /// <summary>
    /// 写入文本。
    /// </summary>
    /// <param name="stream">参与操作的流。</param>
    /// <param name="value">待处理的值。</param>
    private static void WriteText(Stream stream, string value)
    {
        var bytes = Encoding.UTF8.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>
    /// 删除文件。
    /// </summary>
    /// <param name="path">目标文件或目录路径。</param>
    private static void DeleteFile(string path)
    {
        if (File.Exists(path))
            File.Delete(path);

        foreach (var temporaryPath in GetTemporaryFiles(path))
        {
            if (File.Exists(temporaryPath))
                File.Delete(temporaryPath);
        }
    }
}
