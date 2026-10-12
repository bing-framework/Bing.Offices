using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using Bing.Offices;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 前向 XLSX 共享字符串磁盘索引测试。
/// </summary>
public sealed class ExcelForwardXlsxSharedStringsTest
{
    /// <summary>
    /// 验证大量短字符串的索引保存在磁盘且支持随机读取。
    /// </summary>
    [Fact]
    public void ReadDiskBacked_ShouldKeepIndexOnDiskAndSupportRandomReads()
    {
        const int valueCount = 20000;
        var values = new List<string>(valueCount);
        var xml = new StringBuilder("<sst>");
        for (var index = 0; index < valueCount; index++)
        {
            switch (index % 4)
            {
                case 0:
                    values.Add(string.Empty);
                    xml.Append("<si><t></t></si>");
                    break;
                case 1:
                    values.Add($"v{index}");
                    xml.Append("<si><t>v").Append(index).Append("</t></si>");
                    break;
                case 2:
                    values.Add($"多语言{index}");
                    xml.Append("<si><r><t>多</t></r><r><t>语言").Append(index)
                        .Append("</t></r></si>");
                    break;
                default:
                    values.Add($"tail{index}");
                    xml.Append("<si><t>tail").Append(index).Append("</t></si>");
                    break;
            }
        }
        xml.Append("</sst>");

        var directory = CreateTemporaryDirectory();
        try
        {
            using (var package = CreatePackage(xml.ToString()))
            using (var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true))
            using (var strings = ExcelForwardXlsxSharedStrings.ReadDiskBacked(
                archive.GetEntry("xl/sharedStrings.xml"), CancellationToken.None, directory))
            {
                Assert.Equal(valueCount, strings.Count);
                Assert.Equal(values.Sum(value => Encoding.UTF8.GetByteCount(value)),
                    Directory.GetFiles(directory, "bing-offices-forward-shared-text-*.bin")
                        .Select(path => new FileInfo(path).Length).Single());
                Assert.Equal(valueCount * 12L,
                    Directory.GetFiles(directory, "bing-offices-forward-shared-index-*.bin")
                        .Select(path => new FileInfo(path).Length).Single());
                foreach (var index in new[] { 0, 1, 2, 9999, valueCount - 1 })
                    Assert.Equal(values[index], strings[index]);
                Assert.Equal(values, strings.ToArray());
            }

            Assert.Empty(Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证超长共享字符串分块写入时保留 Unicode 内容。
    /// </summary>
    [Fact]
    public void ReadDiskBacked_WhenStringSpansChunks_ShouldPreserveUnicodeValue()
    {
        var value = new string('a', 4095) + "😀" + new string('中', 8192);
        var xml = $"<sst><si><t>{value}</t></si></sst>";
        var directory = CreateTemporaryDirectory();
        try
        {
            using (var package = CreatePackage(xml))
            using (var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true))
            using (var strings = ExcelForwardXlsxSharedStrings.ReadDiskBacked(
                archive.GetEntry("xl/sharedStrings.xml"), CancellationToken.None, directory))
            {
                Assert.Single(strings);
                Assert.Equal(value, strings[0]);
                Assert.Equal(Encoding.UTF8.GetByteCount(value),
                    Directory.GetFiles(directory, "bing-offices-forward-shared-text-*.bin")
                        .Select(path => new FileInfo(path).Length).Single());
            }

            Assert.Empty(Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证解析取消后两个磁盘临时文件均被清理。
    /// </summary>
    [Fact]
    public void ReadDiskBacked_WhenCancelled_ShouldDeleteTemporaryFiles()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            using var package = CreatePackage("<sst><si><t>value</t></si></sst>");
            using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            Assert.ThrowsAny<OperationCanceledException>(() =>
                ExcelForwardXlsxSharedStrings.ReadDiskBacked(
                    archive.GetEntry("xl/sharedStrings.xml"), cancellation.Token, directory));
            Assert.Empty(Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证预取消时空共享字符串表不会跳过取消检查。
    /// </summary>
    [Fact]
    public void ReadDiskBacked_WhenPreCanceledAndEmpty_ShouldThrowBeforeCreatingFiles()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            using var package = CreatePackage("<sst/>");
            using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            Assert.ThrowsAny<OperationCanceledException>(() =>
                ExcelForwardXlsxSharedStrings.ReadDiskBacked(
                    archive.GetEntry("xl/sharedStrings.xml"), cancellation.Token, directory));
            Assert.Empty(Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证 XML 解析失败后两个磁盘临时文件均被清理。
    /// </summary>
    [Fact]
    public void ReadDiskBacked_WhenXmlIsInvalid_ShouldDeleteTemporaryFiles()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            using var package = CreatePackage("<sst><si><t>saved</t></si><si><t>broken</sst>");
            using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);

            Assert.Throws<XmlException>(() => ExcelForwardXlsxSharedStrings.ReadDiskBacked(
                archive.GetEntry("xl/sharedStrings.xml"), CancellationToken.None, directory));
            Assert.Empty(Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static string CreateTemporaryDirectory()
    {
        var directory = Path.Combine(Path.GetTempPath(),
            $"bing-offices-forward-shared-test-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static MemoryStream CreatePackage(string sharedStringsXml)
    {
        var package = new MemoryStream();
        using (var archive = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true))
        using (var writer = new StreamWriter(archive.CreateEntry("xl/sharedStrings.xml").Open(),
                   new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
            writer.Write(sharedStringsXml);
        package.Position = 0;
        return package;
    }
}
