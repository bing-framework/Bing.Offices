using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;

namespace Bing.Offices.Internals;

/// <summary>
/// 为透视模板合并提供按需读取的 ZIP 部件存储。
/// </summary>
/// <remarks>
/// 原始部件保留在可删除的临时 ZIP 文件中，只有被 XML 合并逻辑访问的部件才载入内存；
/// 新增或修改的部件以字节数组保存，最终写出时逐条复制。该类型不接管调用方输入流。
/// </remarks>
internal sealed class ExcelPivotTemplateArchive : IDisposable
{
    private readonly FileStream _file;
    private readonly ZipArchive _archive;
    private readonly Dictionary<string, ZipArchiveEntry> _entries;
    private readonly CancellationToken _token;
    private readonly Dictionary<string, byte[]> _overrides = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _removed = new(StringComparer.OrdinalIgnoreCase);
    private bool _disposed;

    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplateArchive" /> 类型的实例。
    /// </summary>
    /// <param name="file">保存 ZIP 内容的临时文件流。</param>
    /// <param name="archive">已打开的 ZIP 归档。</param>
    /// <param name="entries">按规范化名称索引的原始条目。</param>
    /// <param name="token">归档读取使用的取消令牌。</param>
    internal ExcelPivotTemplateArchive(FileStream file, ZipArchive archive,
        Dictionary<string, ZipArchiveEntry> entries, CancellationToken token)
    {
        _file = file ?? throw new ArgumentNullException(nameof(file));
        _archive = archive ?? throw new ArgumentNullException(nameof(archive));
        _entries = entries ?? throw new ArgumentNullException(nameof(entries));
        _token = token;
    }

    /// <summary>
    /// 获取当前部件名称。
    /// </summary>
    internal IEnumerable<string> Keys
    {
        get
        {
            var names = new HashSet<string>(_entries.Keys, StringComparer.OrdinalIgnoreCase);
            names.UnionWith(_overrides.Keys);
            names.ExceptWith(_removed);
            return names;
        }
    }

    /// <summary>
    /// 获取或设置指定部件的内容。
    /// </summary>
    /// <param name="name">部件名称。</param>
    /// <returns>部件内容。</returns>
    internal byte[] this[string name]
    {
        get => TryGetValue(name, out var bytes)
            ? bytes
            : throw new KeyNotFoundException($"ZIP 部件不存在: {name}");
        set => Set(name, value);
    }

    /// <summary>
    /// 判断是否包含指定部件。
    /// </summary>
    /// <param name="name">部件名称。</param>
    /// <returns>是否存在未删除部件。</returns>
    internal bool ContainsKey(string name)
    {
        name = Normalize(name);
        return !_removed.Contains(name) && (_overrides.ContainsKey(name) || _entries.ContainsKey(name));
    }

    /// <summary>
    /// 按需读取指定部件。
    /// </summary>
    /// <param name="name">部件名称。</param>
    /// <param name="bytes">部件内容。</param>
    /// <returns>是否找到未删除部件。</returns>
    internal bool TryGetValue(string name, out byte[] bytes)
    {
        name = Normalize(name);
        if (_removed.Contains(name))
        {
            bytes = null;
            return false;
        }
        if (_overrides.TryGetValue(name, out bytes))
            return true;
        if (!_entries.TryGetValue(name, out var entry))
        {
            bytes = null;
            return false;
        }
        bytes = ReadEntry(entry, _token);
        return true;
    }

    /// <summary>
    /// 写入新增或修改后的部件。
    /// </summary>
    /// <param name="name">部件名称。</param>
    /// <param name="bytes">部件内容。</param>
    internal void Set(string name, byte[] bytes)
    {
        if (bytes == null)
            throw new ArgumentNullException(nameof(bytes));
        name = Normalize(name);
        _removed.Remove(name);
        _overrides[name] = bytes;
    }

    /// <summary>
    /// 删除指定部件。
    /// </summary>
    /// <param name="name">部件名称。</param>
    /// <returns>是否删除了现有部件。</returns>
    internal bool Remove(string name)
    {
        name = Normalize(name);
        var existed = ContainsKey(name);
        _overrides.Remove(name);
        if (_entries.ContainsKey(name))
            _removed.Add(name);
        return existed;
    }

    /// <summary>
    /// 将当前部件按名称写入目标 ZIP。
    /// </summary>
    /// <param name="target">目标 ZIP。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="compressionLevel">目标压缩级别。</param>
    internal void WriteTo(ZipArchive target, CancellationToken token,
        CompressionLevel compressionLevel)
    {
        if (target == null)
            throw new ArgumentNullException(nameof(target));
        foreach (var name in Keys.OrderBy(item => item, StringComparer.Ordinal))
        {
            token.ThrowIfCancellationRequested();
            var entry = target.CreateEntry(name, compressionLevel);
            using var output = entry.Open();
            if (_overrides.TryGetValue(name, out var bytes))
            {
                WriteBytes(output, bytes, token);
                continue;
            }
            if (!_entries.TryGetValue(name, out var source))
                throw new InvalidDataException($"ZIP 部件不存在: {name}");
            using var input = source.Open();
            Copy(input, output, token);
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _archive.Dispose();
        _file.Dispose();
    }

    /// <summary>
    /// 读取原始 ZIP 条目。
    /// </summary>
    private static byte[] ReadEntry(ZipArchiveEntry entry, CancellationToken token)
    {
        using var input = entry.Open();
        using var output = new MemoryStream();
        var buffer = new byte[81920];
        while (true)
        {
            token.ThrowIfCancellationRequested();
            var count = input.Read(buffer, 0, buffer.Length);
            if (count <= 0)
                break;
            output.Write(buffer, 0, count);
        }
        token.ThrowIfCancellationRequested();
        return output.ToArray();
    }

    /// <summary>
    /// 分块写入内存部件。
    /// </summary>
    private static void WriteBytes(Stream destination, byte[] bytes, CancellationToken token)
    {
        const int bufferSize = 81920;
        var offset = 0;
        while (offset < bytes.Length)
        {
            token.ThrowIfCancellationRequested();
            var count = Math.Min(bufferSize, bytes.Length - offset);
            destination.Write(bytes, offset, count);
            offset += count;
        }
        token.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 分块复制原始部件。
    /// </summary>
    private static void Copy(Stream source, Stream destination, CancellationToken token)
    {
        var buffer = new byte[81920];
        while (true)
        {
            token.ThrowIfCancellationRequested();
            var count = source.Read(buffer, 0, buffer.Length);
            if (count <= 0)
                break;
            destination.Write(buffer, 0, count);
        }
        token.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 规范化 ZIP 部件名称。
    /// </summary>
    private static string Normalize(string name) => name?.Replace('\\', '/').TrimStart('/') ?? string.Empty;
}
