using System.Collections;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Bing.Offices;

/// <summary>
/// 保存 Forward XLSX 共享字符串，并在摘要模式下以临时文件索引降低托管内存占用。
/// </summary>
internal sealed class ExcelForwardXlsxSharedStrings : IReadOnlyList<string>, IDisposable
{
    /// <summary>
    /// 磁盘索引中单条记录的字节数。
    /// </summary>
    private const int IndexRecordSize = sizeof(long) + sizeof(int);

    /// <summary>
    /// UTF-8 编码器。
    /// </summary>
    private static readonly Encoding Utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    /// <summary>
    /// 内存模式共享字符串。
    /// </summary>
    private readonly IReadOnlyList<string> _values;

    /// <summary>
    /// 磁盘索引的临时存储。
    /// </summary>
    private readonly FileStream _storage;

    /// <summary>
    /// 磁盘偏移索引的临时存储。
    /// </summary>
    private readonly FileStream _indexStorage;

    /// <summary>
    /// 读取磁盘偏移索引。
    /// </summary>
    private readonly BinaryReader _indexReader;

    /// <summary>
    /// 当前索引包含的字符串数量。
    /// </summary>
    private readonly int _count;

    /// <summary>
    /// 保护临时存储的定位和读取操作。
    /// </summary>
    private readonly object _storageLock = new object();

    /// <summary>
    /// 初始化内存模式共享字符串。
    /// </summary>
    /// <param name="values">共享字符串。</param>
    private ExcelForwardXlsxSharedStrings(IReadOnlyList<string> values)
    {
        _values = values ?? Array.Empty<string>();
        _count = 0;
    }

    /// <summary>
    /// 初始化磁盘索引模式共享字符串。
    /// </summary>
    /// <param name="storage">字符串临时存储。</param>
    /// <param name="indexStorage">字符串偏移索引临时存储。</param>
    /// <param name="count">共享字符串数量。</param>
    private ExcelForwardXlsxSharedStrings(FileStream storage, FileStream indexStorage, int count)
    {
        _storage = storage ?? throw new ArgumentNullException(nameof(storage));
        _indexStorage = indexStorage ?? throw new ArgumentNullException(nameof(indexStorage));
        _indexReader = new BinaryReader(indexStorage, Utf8, leaveOpen: true);
        _count = count;
        _values = Array.Empty<string>();
    }

    /// <summary>
    /// 获取共享字符串数量。
    /// </summary>
    public int Count => _storage == null ? _values.Count : _count;

    /// <summary>
    /// 获取指定索引的共享字符串。
    /// </summary>
    /// <param name="index">共享字符串索引。</param>
    /// <returns>共享字符串文本。</returns>
    public string this[int index]
    {
        get
        {
            if (index < 0 || index >= Count)
                throw new ArgumentOutOfRangeException(nameof(index));
            if (_storage == null)
                return _values[index] ?? string.Empty;

            byte[] bytes;
            lock (_storageLock)
            {
                _indexStorage.Position = (long)index * IndexRecordSize;
                var offset = _indexReader.ReadInt64();
                var length = _indexReader.ReadInt32();
                if (offset < 0 || length < 0 || offset > _storage.Length
                    || length > _storage.Length - offset)
                    throw new InvalidDataException("Forward 共享字符串磁盘索引无效。");
                if (length == 0)
                    return string.Empty;

                bytes = new byte[length];
                _storage.Position = offset;
                var read = 0;
                while (read < bytes.Length)
                {
                    var count = _storage.Read(bytes, read, bytes.Length - read);
                    if (count == 0)
                        throw new EndOfStreamException("Forward 共享字符串临时索引不完整。");
                    read += count;
                }
            }
            return Utf8.GetString(bytes);
        }
    }

    /// <summary>
    /// 创建内存模式共享字符串。
    /// </summary>
    /// <param name="values">共享字符串。</param>
    /// <returns>共享字符串索引。</returns>
    internal static ExcelForwardXlsxSharedStrings FromValues(IReadOnlyList<string> values)
        => new ExcelForwardXlsxSharedStrings(values);

    /// <summary>
    /// 从 XLSX 共享字符串部件创建磁盘索引。
    /// </summary>
    /// <param name="entry">共享字符串部件。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>磁盘索引共享字符串。</returns>
    internal static ExcelForwardXlsxSharedStrings ReadDiskBacked(ZipArchiveEntry entry,
        CancellationToken cancellationToken)
        => ReadDiskBacked(entry, cancellationToken, Path.GetTempPath());

    /// <summary>
    /// 从 XLSX 共享字符串部件创建磁盘索引。
    /// </summary>
    /// <param name="entry">共享字符串部件。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="temporaryDirectory">临时文件目录。</param>
    /// <returns>磁盘索引共享字符串。</returns>
    internal static ExcelForwardXlsxSharedStrings ReadDiskBacked(ZipArchiveEntry entry,
        CancellationToken cancellationToken, string temporaryDirectory)
    {
        if (entry == null)
            throw new ArgumentNullException(nameof(entry));
        if (string.IsNullOrWhiteSpace(temporaryDirectory))
            throw new ArgumentException("临时目录不能为空。", nameof(temporaryDirectory));
        if (!Directory.Exists(temporaryDirectory))
            throw new DirectoryNotFoundException(temporaryDirectory);
        cancellationToken.ThrowIfCancellationRequested();

        var fileId = Guid.NewGuid().ToString("N");
        var storagePath = Path.Combine(temporaryDirectory,
            $"bing-offices-forward-shared-text-{fileId}.bin");
        var indexPath = Path.Combine(temporaryDirectory,
            $"bing-offices-forward-shared-index-{fileId}.bin");
        FileStream storage = null;
        FileStream indexStorage = null;
        try
        {
            storage = new FileStream(storagePath, FileMode.CreateNew, FileAccess.ReadWrite,
                FileShare.Read, 81920, FileOptions.DeleteOnClose | FileOptions.SequentialScan);
            indexStorage = new FileStream(indexPath, FileMode.CreateNew, FileAccess.ReadWrite,
                FileShare.Read, 81920, FileOptions.DeleteOnClose);
            var valueCount = 0;
            using var indexWriter = new BinaryWriter(indexStorage, Utf8, leaveOpen: true);
            using var stream = entry.Open();
            using var reader = CreateReader(stream);
            var characterBuffer = new char[4096];
            var byteBuffer = new byte[Utf8.GetMaxByteCount(characterBuffer.Length)];
            var encoder = Utf8.GetEncoder();
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType != XmlNodeType.Element
                    || !string.Equals(reader.LocalName, "si", StringComparison.Ordinal))
                    continue;

                var valueOffset = storage.Position;
                var valueLength = 0;
                using (var subtree = reader.ReadSubtree())
                {
                    while (subtree.Read())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (subtree.NodeType == XmlNodeType.Element
                            && string.Equals(subtree.LocalName, "t", StringComparison.Ordinal)
                            && !subtree.IsEmptyElement)
                        {
                            while (subtree.Read())
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                if (subtree.NodeType == XmlNodeType.EndElement
                                    && string.Equals(subtree.LocalName, "t", StringComparison.Ordinal))
                                    break;
                                if (subtree.NodeType != XmlNodeType.Text
                                    && subtree.NodeType != XmlNodeType.CDATA
                                    && subtree.NodeType != XmlNodeType.Whitespace
                                    && subtree.NodeType != XmlNodeType.SignificantWhitespace)
                                    continue;

                                int characterCount;
                                do
                                {
                                    cancellationToken.ThrowIfCancellationRequested();
                                    characterCount = subtree.ReadValueChunk(characterBuffer, 0,
                                        characterBuffer.Length);
                                    if (characterCount == 0)
                                        continue;

                                    var byteCount = encoder.GetBytes(characterBuffer, 0,
                                        characterCount, byteBuffer, 0, flush: false);
                                    if (byteCount > 0)
                                        storage.Write(byteBuffer, 0, byteCount);
                                    valueLength = checked(valueLength + byteCount);
                                } while (characterCount > 0);
                            }
                        }
                    }
                }

                var finalByteCount = encoder.GetBytes(characterBuffer, 0, 0, byteBuffer, 0,
                    flush: true);
                if (finalByteCount > 0)
                    storage.Write(byteBuffer, 0, finalByteCount);
                valueLength = checked(valueLength + finalByteCount);
                indexWriter.Write(valueOffset);
                indexWriter.Write(valueLength);
                valueCount = checked(valueCount + 1);
            }
            indexWriter.Flush();
            storage.Flush();
            indexStorage.Flush();
            storage.Position = 0;
            indexStorage.Position = 0;
            return new ExcelForwardXlsxSharedStrings(storage, indexStorage, valueCount);
        }
        catch
        {
            storage?.Dispose();
            indexStorage?.Dispose();
            TryDelete(storagePath);
            TryDelete(indexPath);
            throw;
        }
    }

    /// <summary>
    /// 获取共享字符串枚举器。
    /// </summary>
    /// <returns>共享字符串枚举器。</returns>
    public IEnumerator<string> GetEnumerator()
    {
        for (var index = 0; index < Count; index++)
            yield return this[index];
    }

    /// <summary>
    /// 获取非泛型共享字符串枚举器。
    /// </summary>
    /// <returns>共享字符串枚举器。</returns>
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>
    /// 释放共享字符串临时索引。
    /// </summary>
    public void Dispose()
    {
        try
        {
            _indexReader?.Dispose();
        }
        finally
        {
            try
            {
                _indexStorage?.Dispose();
            }
            finally
            {
                _storage?.Dispose();
            }
        }
    }

    /// <summary>
    /// 删除创建失败时遗留的临时索引文件。
    /// </summary>
    /// <param name="path">临时文件路径。</param>
    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // 临时文件由 DeleteOnClose 负责兜底清理。
        }
    }

    /// <summary>
    /// 创建安全 XML 读取器。
    /// </summary>
    /// <param name="stream">XML 流。</param>
    /// <returns>XML 读取器。</returns>
    private static XmlReader CreateReader(Stream stream) => XmlReader.Create(stream, new XmlReaderSettings
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreWhitespace = true,
        MaxCharactersFromEntities = 0,
        MaxCharactersInDocument = 0
    });
}
