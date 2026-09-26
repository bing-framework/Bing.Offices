using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证独立 Golden XLSX 资源的 manifest、哈希和资源集合保持一致。
/// </summary>
public sealed class GoldenFixtureManifestTest
{
    /// <summary>
    /// manifest 中的每个文件都必须存在并匹配冻结 SHA-256，且目录不得有未登记输入。
    /// </summary>
    [Fact]
    public void GoldenManifest_ShouldMatchFrozenResources()
    {
        var root = Path.Combine(AppContext.BaseDirectory, "Resources", "Golden");
        var manifestPath = Path.Combine(root, "manifest.json");
        Assert.True(File.Exists(manifestPath), $"Golden manifest was not copied: {manifestPath}");

        var manifest = JsonSerializer.Deserialize<GoldenManifest>(
            File.ReadAllText(manifestPath, Encoding.UTF8), new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
        Assert.NotNull(manifest);
        Assert.Equal("golden-fixtures-v1", manifest.Generator);
        Assert.NotEmpty(manifest.Files);

        var declared = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var file in manifest.Files)
        {
            Assert.False(string.IsNullOrWhiteSpace(file.Path));
            Assert.True(declared.Add(file.Path), $"Duplicate Golden path: {file.Path}");
            Assert.False(string.IsNullOrWhiteSpace(file.Purpose));
            Assert.False(string.IsNullOrWhiteSpace(file.Expected));
            Assert.False(string.IsNullOrWhiteSpace(file.Provenance));
            Assert.False(string.IsNullOrWhiteSpace(file.Generator));
            Assert.False(string.IsNullOrWhiteSpace(file.UpdateReason));
            Assert.Equal("1", file.ContractVersion);
            var path = Path.Combine(root, file.Path.Replace('/', Path.DirectorySeparatorChar));
            Assert.True(File.Exists(path), $"Manifest file is missing: {file.Path}");
            var actual = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
            Assert.Equal(file.Sha256, actual);
        }

        var actualFiles = Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories)
            .Where(path => !string.Equals(Path.GetFileName(path), "manifest.json", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(Path.GetFileName(path), "README.md", StringComparison.OrdinalIgnoreCase))
            .Select(path => Path.GetRelativePath(root, path).Replace(Path.DirectorySeparatorChar, '/'))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Equal(declared.OrderBy(path => path), actualFiles.OrderBy(path => path));
    }

    /// <summary>
    /// 黄金样例资源清单。
    /// </summary>
    private sealed class GoldenManifest
    {
        /// <summary>
        /// 获取或设置资源生成器标识。
        /// </summary>
        public string Generator { get; set; }
        /// <summary>
        /// 获取或设置清单登记的资源集合。
        /// </summary>
        public List<GoldenFile> Files { get; set; }
    }

    /// <summary>
    /// 黄金样例资源的来源与校验记录。
    /// </summary>
    private sealed class GoldenFile
    {
        /// <summary>
        /// 获取或设置资源相对路径。
        /// </summary>
        public string Path { get; set; }
        /// <summary>
        /// 获取或设置资源的测试用途。
        /// </summary>
        public string Purpose { get; set; }
        /// <summary>
        /// 获取或设置资源的预期契约结果。
        /// </summary>
        public string Expected { get; set; }
        /// <summary>
        /// 获取或设置资源来源说明。
        /// </summary>
        public string Provenance { get; set; }
        /// <summary>
        /// 获取或设置资源生成器标识。
        /// </summary>
        public string Generator { get; set; }
        /// <summary>
        /// 获取或设置资源内容的 SHA-256 校验值。
        /// </summary>
        public string Sha256 { get; set; }
        /// <summary>
        /// 获取或设置资源对应的契约版本。
        /// </summary>
        public string ContractVersion { get; set; }
        /// <summary>
        /// 获取或设置资源更新原因。
        /// </summary>
        public string UpdateReason { get; set; }
    }
}
