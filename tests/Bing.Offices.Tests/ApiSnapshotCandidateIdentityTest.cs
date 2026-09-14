using System;
using System.IO;
using System.Text;
using Bing.Offices.ApiSnapshot;
using Xunit;

namespace Bing.Offices.Tests;

public sealed class ApiSnapshotCandidateIdentityTest
{
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

    [Theory]
    [InlineData("")]
    [InlineData("not-a-sha256")]
    [InlineData("000000000000000000000000000000000000000000000000000000000000000g")]
    public void CandidateIdentityHash_RejectsMalformedSha256(string value)
    {
        Assert.False(ApiSnapshotFileHash.IsSha256(value));
    }
}
