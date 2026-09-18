using System;
using System.IO;
using System.Text;
using Bing.Offices.ApiSnapshot;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证候选程序集快照的文件身份校验行为。
/// </summary>
public sealed class ApiSnapshotCandidateIdentityTest
{
    /// <summary>
    /// 验证候选程序集文件被篡改后其哈希发生变化。
    /// </summary>
    [Fact]
    public void CandidateAssemblyHash_ChangesWhenTheCandidateFileIsTampered()
    {
        var path = Path.Combine(Path.GetTempPath(), $"api-snapshot-{Guid.NewGuid():N}.dll");
        try
        {
            File.WriteAllText(path, "candidate", Encoding.UTF8);
            var recordedHash = ApiSnapshotFileHash.ComputeSha256(path);

            File.WriteAllText(path, "tampered candidate", Encoding.UTF8);

            var actualHash = ApiSnapshotFileHash.ComputeSha256(path);
            Assert.True(ApiSnapshotFileHash.IsSha256(recordedHash));
            Assert.NotEqual(recordedHash, actualHash);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 验证格式错误的 SHA-256 哈希会被拒绝。
    /// </summary>
    /// <param name="value">待验证的 SHA-256 哈希字符串。</param>
    [Theory]
    [InlineData("")]
    [InlineData("not-a-sha256")]
    [InlineData("000000000000000000000000000000000000000000000000000000000000000g")]
    public void CandidateIdentityHash_RejectsMalformedSha256(string value)
    {
        Assert.False(ApiSnapshotFileHash.IsSha256(value));
    }
}
