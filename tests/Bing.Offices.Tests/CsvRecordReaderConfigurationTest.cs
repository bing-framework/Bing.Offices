using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Csv;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>直接验证同步／异步 CSV 读取配置、失败恢复与读取器所有权。</summary>
public sealed class CsvRecordReaderConfigurationTest
{
    /// <summary>同步枚举期间插入另一种读取配置，不得污染后续记录。</summary>
    [Fact]
    public async Task Read_InterleavedOptions_ShouldKeepIndependentRecords()
    {
        using var source = new StringReader("\"a;b\";x\r\n\"c;d\";z\r\n");
        using var records = CsvRecordReader.Read(source, ';', '"', default, 3).GetEnumerator();
        Assert.True(records.MoveNext());
        Assert.Equal(new[] { "a;b", "x" }, records.Current);

        using var other = new StringReader("'q|r'|s\r\n");
        AssertRecords(new[] { new[] { "q|r", "s" } }, await Read(other, '|', '\'', null, true));

        Assert.True(records.MoveNext());
        Assert.Equal(new[] { "c;d", "z" }, records.Current);
        Assert.False(records.MoveNext());
        Assert.Equal(-1, source.Peek());
        Assert.Equal(-1, other.Peek());
    }

    /// <summary>异步回调中启动其他配置的读取，原解析器仍使用本次配置。</summary>
    [Fact]
    public async Task ReadAsync_NestedOptions_ShouldKeepIndependentRecords()
    {
        using var source = new StringReader("'a|b'|x\r\n'c|d'|z\r\n");
        var actual = new List<IReadOnlyList<string>>();
        await CsvRecordReader.ReadAsync(source, '|', '\'', async record =>
        {
            actual.Add(record);
            using var nested = new StringReader("\"long;value\";ok\r\n");
            AssertRecords(new[] { new[] { "long;value", "ok" } }, await Read(nested, ';', '"', null, false));
            Assert.Equal(-1, nested.Peek());
        }, default, 3);
        AssertRecords(new[] { new[] { "a|b", "x" }, new[] { "c|d", "z" } }, actual);
        Assert.Equal(-1, source.Peek());
    }

    /// <summary>长度校验失败不保留状态，也不关闭调用方读取器。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Read_LimitFailureThenUnlimitedRead_ShouldNotRetainConfiguration(bool asynchronous)
    {
        using var invalid = new StringReader("ok,abc\r\n");
        var error = await Assert.ThrowsAsync<CsvResourceLimitException>(() => Read(invalid, ',', '"', 2, asynchronous));
        Assert.Equal("CSV 字段超过最大字符数: 2", error.Message);
        Assert.Equal(1, error.RowIndex);
        Assert.Equal(2, error.ColumnIndex);
        Assert.Equal(-1, invalid.Peek());

        using var valid = new StringReader("'long|value'|tail\r\n");
        AssertRecords(new[] { new[] { "long|value", "tail" } }, await Read(valid, '|', '\'', null, asynchronous));
        Assert.Equal(-1, valid.Peek());
    }

    /// <summary>合法的转义、换行和中文均按解码后字符数检查。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Read_QuotedBoundary_ShouldReturnCompleteRecords(bool asynchronous)
    {
        using var reader = new StringReader("\"a\"\"b\",\"x\ny\",中文\r\n");
        AssertRecords(new[] { new[] { "a\"b", "x\ny", "中文" } }, await Read(reader, ',', '"', 3, asynchronous));
        Assert.Equal(-1, reader.Peek());
    }

    /// <summary>非法引用字符之后的数据在两条路径中保留相同异常类型和信息。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Read_MalformedQuotedField_ShouldPreserveBadDataContract(bool asynchronous)
    {
        using var reader = new StringReader("\"value\"x\r\n");
        var error = await Assert.ThrowsAsync<InvalidOperationException>(() => Read(reader, ',', '"', null, asynchronous));
        Assert.Equal("CSV 包含不符合 RFC 4180 的字段。", error.Message);
        Assert.Equal(-1, reader.Peek());
    }

    /// <summary>创建同步枚举器不触发读取，真正枚举时才访问调用方资源。</summary>
    [Fact]
    public void Read_ShouldRemainDeferred()
    {
        using var reader = new ThrowingReader();
        var records = CsvRecordReader.Read(reader, ',', '"', CancellationToken.None);
        Assert.Equal(0, reader.ReadCalls);
        using var iterator = records.GetEnumerator();
        Assert.Equal(0, reader.ReadCalls);
        Assert.Throws<IOException>(() => iterator.MoveNext());
        Assert.Equal(1, reader.ReadCalls);
    }

    private static async Task<IReadOnlyList<IReadOnlyList<string>>> Read(TextReader reader, char delimiter,
        char quote, int? limit, bool asynchronous)
    {
        if (!asynchronous)
            return CsvRecordReader.Read(reader, delimiter, quote, default, limit).ToList();
        var result = new List<IReadOnlyList<string>>();
        await CsvRecordReader.ReadAsync(reader, delimiter, quote, record =>
        {
            result.Add(record);
            return Task.CompletedTask;
        }, default, limit);
        return result;
    }

    private static void AssertRecords(string[][] expected, IReadOnlyList<IReadOnlyList<string>> actual)
    {
        Assert.Equal(expected.Length, actual.Count);
        for (var index = 0; index < expected.Length; index++)
            Assert.Equal(expected[index], actual[index]);
    }

    private sealed class ThrowingReader : TextReader
    {
        public int ReadCalls { get; private set; }

        public override int Read(char[] buffer, int index, int count)
        {
            ReadCalls++;
            throw new IOException("Deferred read");
        }
    }
}
