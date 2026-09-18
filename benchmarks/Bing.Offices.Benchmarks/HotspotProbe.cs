using System.Diagnostics;
using System.Collections;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Xml;
using Bing.Offices.Imports;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 采集 RawDateReader 与关系绑定的 before/after 原始样本。
/// </summary>
internal static class HotspotProbe
{
    private static readonly string[] RawColumnWorkloads = { "3", "10", "30" };
    private static readonly string[] RelationWorkloads = { "1000", "10000", "100000" };

    public static async Task RunAsync(string artifactPath, string phase, int repetitions)
    {
        if (!string.Equals(phase, "before", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(phase, "after", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("phase 必须是 before 或 after。", nameof(phase));
        if (repetitions < 3)
            throw new ArgumentOutOfRangeException(nameof(repetitions), "性能探针至少需要三次重复。");

        var fullPath = Path.GetFullPath(artifactPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        using var writer = new StreamWriter(fullPath, false, new UTF8Encoding(false));
        writer.WriteLine(JsonSerializer.Serialize(new
        {
            kind = "hotspot-probe",
            schema = 1,
            generatedUtc = DateTimeOffset.UtcNow,
            phase,
            repetitions,
            rawDateRows = 100000,
            rawDateColumns = RawColumnWorkloads,
            relationRows = RelationWorkloads,
            beforeSource = "investigative numeric-cell replay; not a HEAD Reader baseline",
            candidateIdentity = GetCandidateIdentity()
        }));

        foreach (var columns in RawColumnWorkloads)
            await RunScenarioAsync(writer, phase, repetitions, "raw-date", int.Parse(columns))
                .ConfigureAwait(false);
        foreach (var rows in RelationWorkloads)
            await RunScenarioAsync(writer, phase, repetitions, "relation", int.Parse(rows))
                .ConfigureAwait(false);

        Console.WriteLine($"HOTSPOT_PROBE artifact={fullPath} phase={phase} "
            + $"rawDateScenarios={RawColumnWorkloads.Length} relationScenarios={RelationWorkloads.Length} status=passed");
    }

    public static void RunWorker(string artifactPath, string phase, string workload, int size,
        int repetitions)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(artifactPath))!);
        using var writer = new StreamWriter(artifactPath, false, new UTF8Encoding(false));
        if (workload == "raw-date")
        {
            var workbook = CreateWorkbook(100000, size);
            MeasureRawDate(workbook, phase, size);
            for (var repetition = 1; repetition <= repetitions; repetition++)
            {
                var sample = MeasureRawDate(workbook, phase, size);
                writer.WriteLine(JsonSerializer.Serialize(sample with { repetition = repetition }));
                writer.Flush();
            }
            return;
        }
        if (workload == "relation")
        {
            var request = CreateRelationRequest();
            BindRelations(request, phase, size);
            for (var repetition = 1; repetition <= repetitions; repetition++)
            {
                var sample = BindRelations(request, phase, size);
                writer.WriteLine(JsonSerializer.Serialize(sample with { repetition = repetition }));
                writer.Flush();
            }
            return;
        }
        throw new ArgumentException($"未知 hotspot workload: {workload}", nameof(workload));
    }

    private static async Task RunScenarioAsync(StreamWriter writer, string phase, int repetitions,
        string workload, int size)
    {
        var path = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(writer.BaseStream is FileStream file
            ? file.Name : throw new InvalidOperationException("探针输出必须是文件。")))!,
            $".{Guid.NewGuid():N}.{workload}.{size}.jsonl");
        try
        {
            await RunWorkerProcessAsync(path, phase, workload, size, repetitions).ConfigureAwait(false);
            var samples = File.ReadLines(path, Encoding.UTF8)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => JsonSerializer.Deserialize<HotspotSample>(line)!)
                .ToArray();
            if (samples.Length != repetitions)
                throw new InvalidOperationException($"{workload}/{size} 样本数不正确: {samples.Length}");
            foreach (var sample in samples)
                writer.WriteLine(JsonSerializer.Serialize(sample));
            writer.WriteLine(JsonSerializer.Serialize(new
            {
                kind = "hotspot-summary",
                workload,
                size,
                phase,
                repetitions,
                medianElapsedMilliseconds = Median(samples.Select(sample => sample.elapsedMilliseconds)),
                medianAllocatedBytes = Median(samples.Select(sample => sample.allocatedBytes)),
                medianPeakWorkingSetBytes = Median(samples.Select(sample => sample.peakWorkingSetBytes)),
                medianIndexCount = Median(samples.Select(sample => (long)sample.indexCount)),
                medianLohSnapshotBytes = Median(samples.Select(sample => sample.lohSnapshotBytes))
            }));
            writer.Flush();
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    private static async Task RunWorkerProcessAsync(string artifactPath, string phase, string workload,
        int size, int repetitions)
    {
        var processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("无法定位性能探针进程。");
        var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = processPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
        };
        if (Path.GetFileNameWithoutExtension(processPath).Equals("dotnet", StringComparison.OrdinalIgnoreCase))
            process.StartInfo.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
        process.StartInfo.ArgumentList.Add("--hotspot-worker");
        process.StartInfo.ArgumentList.Add(artifactPath);
        process.StartInfo.ArgumentList.Add(phase);
        process.StartInfo.ArgumentList.Add(workload);
        process.StartInfo.ArgumentList.Add(size.ToString());
        process.StartInfo.ArgumentList.Add(repetitions.ToString());
        using (process)
        {
            if (!process.Start())
                throw new InvalidOperationException("无法启动 hotspot worker。");
            var stderrTask = process.StandardError.ReadToEndAsync();
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            await process.WaitForExitAsync().ConfigureAwait(false);
            var stderr = await stderrTask.ConfigureAwait(false);
            var stdout = await stdoutTask.ConfigureAwait(false);
            if (process.ExitCode != 0)
                throw new InvalidOperationException($"hotspot worker 失败: {stderr}{stdout}");
        }
    }

    private static HotspotSample MeasureRawDate(byte[] workbook, string phase, int columns)
    {
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        var process = Process.GetCurrentProcess();
        process.Refresh();
        var allocatedBefore = GC.GetTotalAllocatedBytes(true);
        var gen0Before = GC.CollectionCount(0);
        var gen1Before = GC.CollectionCount(1);
        var gen2Before = GC.CollectionCount(2);
        var lohBefore = GetLohSnapshotBytes();
        var stopwatch = Stopwatch.StartNew();
        int indexCount;
        using (var source = new MemoryStream(workbook, writable: false))
        {
            indexCount = string.Equals(phase, "before", StringComparison.OrdinalIgnoreCase)
                ? ReadAllNumeric(source)
                : ReadCurrentDateColumns(source);
        }
        stopwatch.Stop();
        process.Refresh();
        var lohAfter = GetLohSnapshotBytes();
        return new HotspotSample("hotspot-sample", "raw-date", columns, phase, 0,
            stopwatch.Elapsed.TotalMilliseconds,
            GC.GetTotalAllocatedBytes(true) - allocatedBefore, indexCount,
            workbook.Length, GC.CollectionCount(0) - gen0Before,
            GC.CollectionCount(1) - gen1Before, GC.CollectionCount(2) - gen2Before,
            Math.Max(lohBefore, lohAfter), process.PeakWorkingSet64, Environment.ProcessId);
    }

    private static int ReadCurrentDateColumns(Stream source)
    {
        var type = typeof(Bing.Offices.Imports.MiniExcelExcelImporter).Assembly
            .GetType("Bing.Offices.MiniExcel.Internals.MiniExcelRawDateSerialReader")!;
        var method = type.GetMethod("Read", BindingFlags.Static | BindingFlags.NonPublic,
            binder: null, new[] { typeof(Stream), typeof(string), typeof(CancellationToken) },
            modifiers: null)
            ?? throw new MissingMethodException(type.FullName, "Read(Stream, string, CancellationToken)");
        var result = (IReadOnlyDictionary<long, double>)method.Invoke(null,
            new object[] { source, "Data", CancellationToken.None })!;
        return result.Count;
    }

    private static int ReadAllNumeric(Stream source)
    {
        source.Position = 0;
        using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        using var stream = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true
        });
        var values = new Dictionary<long, double>();
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c")
                continue;
            var reference = reader.GetAttribute("r");
            var value = ReadCellValue(reader);
            if (TryParseReference(reference, out var row, out var column)
                && double.TryParse(value, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var serial))
                values[((long)row << 32) | (uint)column] = serial;
        }
        return values.Count;
    }

    private static string ReadCellValue(XmlReader reader)
    {
        using var subtree = reader.ReadSubtree();
        while (subtree.Read())
            if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "v")
                return subtree.ReadElementContentAsString();
        return string.Empty;
    }

    private static HotspotSample BindRelations(ExcelWorkbookImportRequest<RelationProbeWorkbook> request,
        string phase, int size)
    {
        var root = CreateRelationWorkbook(size);
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        var process = Process.GetCurrentProcess();
        process.Refresh();
        var allocatedBefore = GC.GetTotalAllocatedBytes(true);
        var gen0Before = GC.CollectionCount(0);
        var gen1Before = GC.CollectionCount(1);
        var gen2Before = GC.CollectionCount(2);
        var lohBefore = GetLohSnapshotBytes();
        var stopwatch = Stopwatch.StartNew();
        if (string.Equals(phase, "before", StringComparison.OrdinalIgnoreCase))
        {
            var relation = request.Relations.Single();
            var parents = ((IEnumerable)relation.Parents(root)).Cast<object>().ToArray();
            var children = ((IEnumerable)relation.Children(root)).Cast<object>().ToArray();
            foreach (var child in children)
            {
                var key = relation.ChildKey.DynamicInvoke(child);
                var parent = parents.FirstOrDefault(candidate => RelationKeysEqual(
                    relation.ParentKey.DynamicInvoke(candidate), key, relation.Comparer));
                (relation.Navigation(parent) as IList)?.Add(child);
            }
        }
        else
        {
            var method = typeof(Bing.Offices.Imports.MiniExcelExcelImporter)
                .GetMethod("BindRelations", BindingFlags.Static | BindingFlags.NonPublic)!
                .MakeGenericMethod(typeof(RelationProbeWorkbook));
            var errors = new List<ExcelImportError>();
            method.Invoke(null, new object[] { root, request.Relations, errors, request, CancellationToken.None });
        }
        stopwatch.Stop();
        process.Refresh();
        var lohAfter = GetLohSnapshotBytes();
        var linked = root.Parents.Sum(parent => parent.Items.Count);
        if (linked != size)
            throw new InvalidOperationException($"关系探针绑定数量错误: {linked}/{size}");
        return new HotspotSample("hotspot-sample", "relation", size, phase, 0,
            stopwatch.Elapsed.TotalMilliseconds,
            GC.GetTotalAllocatedBytes(true) - allocatedBefore, linked, 0,
            GC.CollectionCount(0) - gen0Before, GC.CollectionCount(1) - gen1Before,
            GC.CollectionCount(2) - gen2Before, Math.Max(lohBefore, lohAfter),
            process.PeakWorkingSet64, Environment.ProcessId);
    }

    private static bool RelationKeysEqual(object? left, object? right, object? comparer)
    {
        if (comparer is System.Collections.IEqualityComparer nonGeneric)
            return nonGeneric.Equals(left, right);
        if (comparer is IEqualityComparer<object> objectComparer)
            return objectComparer.Equals(left, right);
        if (comparer != null)
        {
            var leftType = left?.GetType() ?? right?.GetType();
            if (leftType != null)
            {
                var equals = comparer.GetType().GetMethod(nameof(object.Equals),
                    BindingFlags.Instance | BindingFlags.Public, binder: null,
                    types: new[] { leftType, leftType }, modifiers: null);
                if (equals != null)
                    return equals.Invoke(comparer, new[] { left, right }) is true;
            }
        }
        return Equals(left, right);
    }

    private static long GetLohSnapshotBytes() =>
        GC.GetGCMemoryInfo().GenerationInfo.Length > 3
            ? GC.GetGCMemoryInfo().GenerationInfo[3].SizeAfterBytes : 0;

    private static ExcelWorkbookImportRequest<RelationProbeWorkbook> CreateRelationRequest() =>
        ExcelImport.Workbook<RelationProbeWorkbook>(workbook =>
        {
            workbook.Sheet("Probe", root => root.Parents);
            workbook.HasMany(root => root.Parents, root => root.Children,
                parent => parent.Key, child => child.Key,
                parent => parent.Items, StringComparer.OrdinalIgnoreCase);
        });

    private static RelationProbeWorkbook CreateRelationWorkbook(int size)
    {
        var root = new RelationProbeWorkbook();
        for (var index = 0; index < size; index++)
        {
            root.Parents.Add(new RelationProbeParent { Key = $"K-{index}" });
            root.Children.Add(new RelationProbeChild { Key = $"k-{index}" });
        }
        return root;
    }

    private static byte[] CreateWorkbook(int rows, int columns)
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "xl/workbook.xml",
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>"
                + "<sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            WriteEntry(archive, "xl/_rels/workbook.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                + "</Relationships>");
            var entry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
            using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false), 65536);
            writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
            for (var row = 1; row <= rows; row++)
            {
                writer.Write($"<row r=\"{row}\">");
                for (var column = 1; column <= columns; column++)
                    writer.Write($"<c r=\"{ColumnName(column)}{row}\" t=\"n\"><v>{row + column}</v></c>");
                writer.Write("</row>");
            }
            writer.Write("</sheetData></worksheet>");
        }
        return stream.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string name, string value)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(value);
    }

    private static string ColumnName(int column)
    {
        var result = string.Empty;
        while (column > 0)
        {
            column--;
            result = (char)('A' + column % 26) + result;
            column /= 26;
        }
        return result;
    }

    private static bool TryParseReference(string? reference, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (string.IsNullOrWhiteSpace(reference))
            return false;
        var index = 0;
        long parsedColumn = 0;
        while (index < reference.Length && char.IsLetter(reference[index]))
        {
            parsedColumn = parsedColumn * 26 + char.ToUpperInvariant(reference[index]) - 'A' + 1;
            index++;
        }
        if (index == 0 || index >= reference.Length
            || !int.TryParse(reference.Substring(index), out row) || row <= 0)
            return false;
        column = (int)parsedColumn;
        return column > 0;
    }

    private static IReadOnlyDictionary<string, string> GetCandidateIdentity()
    {
        var types = new[]
        {
            typeof(HotspotProbe),
            typeof(Bing.Offices.Exports.IExcelExporter),
            typeof(Bing.Offices.Providers.IExcelMappingPlanFactory),
            typeof(Bing.Offices.IO.DefaultFileExportCommitter),
            typeof(Bing.Offices.Imports.MiniExcelExcelImporter)
        };
        return types.Select(type => type.Assembly.Location)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToDictionary(path => Path.GetFileName(path),
                path => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(
                    File.ReadAllBytes(path))), StringComparer.Ordinal);
    }

    private static double Median(IEnumerable<double> values)
    {
        var ordered = values.OrderBy(value => value).ToArray();
        return ordered.Length == 0 ? 0 : ordered[ordered.Length / 2];
    }

    private static double Median(IEnumerable<long> values) => Median(values.Select(value => (double)value));

    private sealed record HotspotSample(string kind, string workload, int size, string phase, int repetition,
        double elapsedMilliseconds, long allocatedBytes, int indexCount, int workbookBytes,
        int gen0Collections, int gen1Collections, int gen2Collections, long lohSnapshotBytes,
        long peakWorkingSetBytes, int processId);

    private sealed class RelationProbeWorkbook
    {
        public List<RelationProbeParent> Parents { get; } = new();
        public List<RelationProbeChild> Children { get; } = new();
    }

    private sealed class RelationProbeParent
    {
        public string Key { get; set; } = string.Empty;
        public List<RelationProbeChild> Items { get; } = new();
    }

    private sealed class RelationProbeChild
    {
        public string Key { get; set; } = string.Empty;
    }
}
