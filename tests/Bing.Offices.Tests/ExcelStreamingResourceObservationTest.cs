#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>验证流式容量探针对本次原子提交文件的范围和峰值统计。</summary>
public class ExcelStreamingResourceObservationTest
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExportCommitter_SamplesTheSessionAndCommitCopyOnEveryWrite(bool useAsync)
    {
        var root = Path.Combine(Path.GetTempPath(), "bing-resource-commit-" + Guid.NewGuid().ToString("N"));
        var sessionDirectory = Path.Combine(root, "session");
        Directory.CreateDirectory(sessionDirectory);
        var targetPath = Path.Combine(root, "result.xlsx");
        var commitFiles = new global::ExcelStreamingResourceMatrix.CommitTemporaryFileTracker(targetPath);
        var npoiDirectory = Path.Combine(Path.GetTempPath(), "poifiles");
        var npoiBaseline = new HashSet<string>(Directory.Exists(npoiDirectory)
            ? Directory.GetFiles(npoiDirectory, "*", SearchOption.AllDirectories)
            : Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);

        try
        {
            using var sampler = new global::ExcelStreamingResourceMatrix.ResourceSampler(sessionDirectory,
                npoiDirectory, npoiBaseline, commitFiles);
            var serviceCollection = new ServiceCollection();
            serviceCollection.AddSingleton<IFileExportCommitter>(
                new global::ExcelStreamingResourceMatrix.SamplingFileExportCommitter(
                    new DefaultFileExportCommitter(), sampler.SampleAfterCommitWrite));
            serviceCollection.AddBingOfficesNpoiStreaming();
            using var services = serviceCollection.BuildServiceProvider();

            var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", new[]
            {
                new CommitProbeRow { Value = "first" },
                new CommitProbeRow { Value = "second" }
            }));

            sampler.Start();
            var exporter = services.GetRequiredService<IExcelStreamingExporter>();
            if (useAsync)
                await exporter.ExportToFileAsync(request, targetPath,
                    new() { TempDirectory = sessionDirectory }).ConfigureAwait(false);
            else
                exporter.ExportToFile(request, targetPath, new() { TempDirectory = sessionDirectory });

            var snapshot = sampler.Stop();
            Assert.True(snapshot.SessionAndCommitOverlapObserved);
            Assert.True(snapshot.LibrarySessionTempPeak > 0);
            Assert.True(snapshot.CommitTempPeak > 0);
            Assert.True(File.Exists(targetPath));
            var targetLength = new FileInfo(targetPath).Length;
            Assert.True(targetLength > 0);
            Assert.Equal(targetLength, snapshot.CommitTempPeak);
            Assert.True(snapshot.LibraryTempPeak >=
                        snapshot.LibrarySessionTempPeak + snapshot.CommitTempPeak);
            Assert.Equal(0, commitFiles.GetCurrentCount());
        }
        finally
        {
            if (Directory.Exists(root))
                Directory.Delete(root, true);
        }
    }

    [Fact]
    public void ResourceSampler_IncludesOnlyCurrentTargetCommitFile_WithSessionBytes()
    {
        var root = Path.Combine(Path.GetTempPath(), "bing-resource-observation-" + Guid.NewGuid().ToString("N"));
        var sessionDirectory = Path.Combine(root, "session");
        var npoiDirectory = Path.Combine(root, "poifiles");
        Directory.CreateDirectory(sessionDirectory);
        Directory.CreateDirectory(npoiDirectory);

        var targetPath = Path.Combine(root, "result.xlsx");
        var previousCommit = targetPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        var otherTargetCommit = Path.Combine(root, "other.xlsx." + Guid.NewGuid().ToString("N") + ".tmp");
        var malformedCommit = targetPath + ".not-a-guid.tmp";
        var nearMatchCommit = targetPath + ".tmp";
        var activeCommit = targetPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            File.WriteAllBytes(targetPath, new byte[29]);
            File.WriteAllBytes(previousCommit, new byte[23]);
            File.WriteAllBytes(otherTargetCommit, new byte[31]);
            File.WriteAllBytes(malformedCommit, new byte[37]);
            File.WriteAllBytes(nearMatchCommit, new byte[41]);
            var tracker = new global::ExcelStreamingResourceMatrix.CommitTemporaryFileTracker(targetPath);

            File.WriteAllBytes(Path.Combine(sessionDirectory, "package.tmp"), new byte[11]);
            using (var commitWindow = new FileStream(activeCommit, FileMode.CreateNew, FileAccess.Write,
                       FileShare.ReadWrite))
            {
                commitWindow.Write(new byte[17], 0, 17);
                commitWindow.Flush();

                using var sampler = new global::ExcelStreamingResourceMatrix.ResourceSampler(sessionDirectory,
                    npoiDirectory, new HashSet<string>(StringComparer.OrdinalIgnoreCase), tracker);
                sampler.Start();
                var snapshot = sampler.Stop();

                Assert.Equal(11, snapshot.LibrarySessionTempPeak);
                Assert.Equal(17, snapshot.CommitTempPeak);
                Assert.Equal(28, snapshot.LibraryTempPeak);
                Assert.Equal(1, tracker.GetCurrentCount());
                Assert.Equal(17, tracker.GetCurrentBytes());
            }

            File.Delete(activeCommit);
            Assert.Equal(0, tracker.GetCurrentCount());
        }
        finally
        {
            if (Directory.Exists(root))
                Directory.Delete(root, true);
        }
    }

    private sealed class CommitProbeRow
    {
        public string Value { get; set; }
    }
}
#endif
