using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

internal static class ExcelStreamingResourceMatrix
{
    private const int ColumnCount = 10;
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string DocumentRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ContentTypeNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    internal static async Task<int> RunAsync(string outputDirectory, string direction, string kind, string rowsValue)
    {
        if (!int.TryParse(rowsValue, NumberStyles.None, CultureInfo.InvariantCulture, out var rows)
            || rows is not 10000 and not 100000
            || !IsOneOf(direction, "export", "import")
            || !IsOneOf(kind, "repeated", "unique"))
            return 2;
        direction = direction.ToLowerInvariant();
        kind = kind.ToLowerInvariant();

        outputDirectory = Path.GetFullPath(outputDirectory);
        Directory.CreateDirectory(outputDirectory);
        Console.OutputEncoding = new UTF8Encoding(false);

        var reportPath = Path.Combine(outputDirectory, $"streaming-{direction}-{kind}-{rows}.json");
        var workbookPath = Path.Combine(outputDirectory, $"streaming-{direction}-{kind}-{rows}.xlsx");
        var tempDirectory = Path.Combine(outputDirectory, ".streaming-temp-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDirectory);
        var commitTemporaryFiles = new CommitTemporaryFileTracker(workbookPath);
        var npoiDirectory = Path.Combine(Path.GetTempPath(), "poifiles");
        var npoiFileBaseline = new HashSet<string>(GetFiles(npoiDirectory), StringComparer.OrdinalIgnoreCase);

        var observation = new ScenarioObservation();
        var status = "success";
        string error = null;
        double elapsedMilliseconds = 0;
        ResourceSnapshot resourceSnapshot = default;
        long processPeakWorkingSet = 0;
        var spillObserved = false;
        var report = new ExcelStreamingScenarioReport();

        try
        {
            if (direction == "import")
                CreateSharedStringFixture(workbookPath, kind, rows);

            using var sampler = new ResourceSampler(tempDirectory, npoiDirectory, npoiFileBaseline,
                commitTemporaryFiles);
            var servicesCollection = new ServiceCollection();
            servicesCollection.AddSingleton<IFileExportCommitter>(new SamplingFileExportCommitter(
                new DefaultFileExportCommitter(), sampler.SampleAfterCommitWrite));
            servicesCollection.AddBingOfficesNpoiStreaming();
            using var services = servicesCollection.BuildServiceProvider();

            var exportRequest = direction == "export"
                ? ExcelExport.Workbook(builder => builder.AddSheet("Data", CreateRows(kind, rows, observation)))
                : null;
            var importRequest = direction == "import" ? ExcelStreamingImport.Sheet<StreamingProbeRow>("Data") : null;
            var exporter = direction == "export" ? services.GetRequiredService<IExcelStreamingExporter>() : null;
            var importer = direction == "import" ? services.GetRequiredService<IExcelStreamingImporter>() : null;

            sampler.Start();
            var startedTimestamp = Stopwatch.GetTimestamp();
            var stopwatch = Stopwatch.StartNew();
            try
            {
                if (direction == "export")
                {
                    await exporter.ExportToFileAsync(exportRequest, workbookPath,
                        new ExcelStreamingExportOptions { TempDirectory = tempDirectory }).ConfigureAwait(false);
                }
                else
                {
                    var importOptions = new ExcelStreamingImportOptions { TempDirectory = tempDirectory };
                    using var input = File.OpenRead(workbookPath);
                    long importedRows = 0;
                    var summary = await importer.ImportAsync(importRequest, input, (row, token) =>
                    {
                        token.ThrowIfCancellationRequested();
                        if (row.RowIndex == 1)
                        {
                            observation.MarkFirstRow();
                            spillObserved = ObserveSharedStringSpill(tempDirectory, rows);
                        }
                        if (!row.IsValid)
                            throw new InvalidDataException($"第 {row.RowIndex} 行未成功绑定: {string.Join(";", row.Errors.Select(item => item.Message))}");
                        if (row.RowIndex != importedRows + 1)
                            throw new InvalidDataException($"行号不连续: expected={importedRows + 1}, actual={row.RowIndex}");
                        ValidateRow(row.Item, kind, checked((int)importedRows));
                        importedRows++;
                        return Task.FromResult(ExcelRowReadDecision.Continue);
                    }, importOptions).ConfigureAwait(false);
                    if (!summary.IsSuccess || summary.RowsRead != rows || summary.SuccessfulRows != rows
                        || summary.DeliveredRows != rows || importedRows != rows)
                        throw new InvalidDataException($"流式导入统计不符: rowsRead={summary.RowsRead}, successful={summary.SuccessfulRows}, delivered={summary.DeliveredRows}");
                }
            }
            catch (Exception exception)
            {
                status = "api-failed";
                error = exception.GetType().FullName + ": " + exception.Message;
            }
            finally
            {
                stopwatch.Stop();
                elapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds;
                resourceSnapshot = sampler.Stop();
                processPeakWorkingSet = sampler.ProcessPeakWorkingSet;
            }

            if (status == "success")
            {
                try
                {
                    if (direction == "export")
                        ValidateExportedWorkbook(workbookPath, kind, rows);
                    else if (kind == "unique" && !spillObserved)
                        throw new InvalidDataException("高唯一值导入没有观察到输入、共享字符串数据和索引三个临时文件。");
                }
                catch (Exception exception)
                {
                    status = "verification-failed";
                    error = exception.GetType().FullName + ": " + exception.Message;
                }
            }

            var librarySessionTemporaryFiles = CountFiles(tempDirectory);
            var libraryCommitTemporaryFiles = commitTemporaryFiles.GetCurrentCount();
            var libraryTemporaryFiles = librarySessionTemporaryFiles + libraryCommitTemporaryFiles;
            var npoiTemporaryFiles = CountNewFiles(npoiDirectory, npoiFileBaseline);
            if (libraryTemporaryFiles > 0 && status == "success")
            {
                status = "cleanup-leak";
                error = "API 返回后库自有临时文件仍有残留。";
            }
            spillObserved = direction == "import" && kind == "unique" && spillObserved;
            var firstRowMs = observation.FirstRowTimestamp == 0
                ? (double?)null
                : Stopwatch.GetElapsedTime(startedTimestamp, observation.FirstRowTimestamp).TotalMilliseconds;
            var rowsPerSecond = elapsedMilliseconds <= 0 ? 0 : rows * 1000d / elapsedMilliseconds;
            report = new ExcelStreamingScenarioReport
            {
                Rows = rows,
                Columns = ColumnCount,
                Kind = kind,
                Direction = direction,
                ElapsedMs = elapsedMilliseconds,
                RowsPerSecond = rowsPerSecond,
                FirstRowMs = firstRowMs,
                PeakWorkingSet = processPeakWorkingSet,
                SampledPeakWorkingSet = resourceSnapshot.SampledPeakWorkingSet,
                PeakManagedBytes = resourceSnapshot.PeakManagedBytes,
                LibraryTempPeak = resourceSnapshot.LibraryTempPeak,
                LibrarySessionTempPeak = resourceSnapshot.LibrarySessionTempPeak,
                CommitTempPeak = resourceSnapshot.CommitTempPeak,
                NpoiTempPeak = resourceSnapshot.NpoiTempPeak,
                CommitOverlapObserved = resourceSnapshot.SessionAndCommitOverlapObserved,
                SpillObserved = spillObserved,
                TempFilesRemaining = libraryTemporaryFiles + npoiTemporaryFiles,
                LibraryTempFilesRemaining = libraryTemporaryFiles,
                LibrarySessionTempFilesRemaining = librarySessionTemporaryFiles,
                CommitTempFilesRemaining = libraryCommitTemporaryFiles,
                NpoiTempFilesRemaining = npoiTemporaryFiles,
                Status = status,
                Runtime = RuntimeInformation.FrameworkDescription,
                OperatingSystem = RuntimeInformation.OSDescription,
                OSArchitecture = RuntimeInformation.OSArchitecture.ToString(),
                ProcessArchitecture = RuntimeInformation.ProcessArchitecture.ToString(),
                ProcessorCount = Environment.ProcessorCount,
                Is64BitProcess = Environment.Is64BitProcess,
                MeasurementNotes = "LibraryTempPeak 为每次采样时 session 临时目录与本次目标提交 .tmp 的字节总和；CommitTempPeak 单独记录提交文件。临时盘使用 20ms 定时采样，并在提交输出流每次 Write/WriteAsync 完成后即时采样；该采样以活动输出流的 Length 补足尚未反映到文件元数据的缓冲字节，不调用 Flush。其余峰值为 20ms 采样值。PeakWorkingSet 使用 Process.PeakWorkingSet64，包含进程初始化。导入 fixture 创建和导出结果全量读回均在 API 计时区间之外。",
                Error = error
            };
        }
        catch (Exception exception)
        {
            status = "setup-failed";
            error = exception.GetType().FullName + ": " + exception.Message;
            var librarySessionTemporaryFiles = CountFiles(tempDirectory);
            var libraryCommitTemporaryFiles = commitTemporaryFiles.GetCurrentCount();
            var libraryTemporaryFiles = librarySessionTemporaryFiles + libraryCommitTemporaryFiles;
            var npoiTemporaryFiles = CountNewFiles(npoiDirectory, npoiFileBaseline);
            report = new ExcelStreamingScenarioReport
            {
                Rows = rows,
                Columns = ColumnCount,
                Kind = kind,
                Direction = direction,
                ElapsedMs = elapsedMilliseconds,
                RowsPerSecond = elapsedMilliseconds <= 0 ? 0 : rows * 1000d / elapsedMilliseconds,
                FirstRowMs = observation.FirstRowTimestamp == 0 ? null : Stopwatch.GetElapsedTime(observation.FirstRowTimestamp).TotalMilliseconds,
                PeakWorkingSet = processPeakWorkingSet,
                SampledPeakWorkingSet = resourceSnapshot.SampledPeakWorkingSet,
                PeakManagedBytes = resourceSnapshot.PeakManagedBytes,
                LibraryTempPeak = resourceSnapshot.LibraryTempPeak,
                LibrarySessionTempPeak = resourceSnapshot.LibrarySessionTempPeak,
                CommitTempPeak = resourceSnapshot.CommitTempPeak,
                NpoiTempPeak = resourceSnapshot.NpoiTempPeak,
                CommitOverlapObserved = resourceSnapshot.SessionAndCommitOverlapObserved,
                SpillObserved = spillObserved,
                TempFilesRemaining = libraryTemporaryFiles + npoiTemporaryFiles,
                LibraryTempFilesRemaining = libraryTemporaryFiles,
                LibrarySessionTempFilesRemaining = librarySessionTemporaryFiles,
                CommitTempFilesRemaining = libraryCommitTemporaryFiles,
                NpoiTempFilesRemaining = npoiTemporaryFiles,
                Status = status,
                Runtime = RuntimeInformation.FrameworkDescription,
                OperatingSystem = RuntimeInformation.OSDescription,
                OSArchitecture = RuntimeInformation.OSArchitecture.ToString(),
                ProcessArchitecture = RuntimeInformation.ProcessArchitecture.ToString(),
                ProcessorCount = Environment.ProcessorCount,
                Is64BitProcess = Environment.Is64BitProcess,
                MeasurementNotes = "LibraryTempPeak 为每次采样时 session 临时目录与本次目标提交 .tmp 的字节总和；CommitTempPeak 单独记录提交文件。临时盘使用 20ms 定时采样，并在提交输出流每次 Write/WriteAsync 完成后即时采样；该采样以活动输出流的 Length 补足尚未反映到文件元数据的缓冲字节，不调用 Flush。其余峰值为 20ms 采样值。PeakWorkingSet 使用 Process.PeakWorkingSet64，包含进程初始化。导入 fixture 创建和导出结果全量读回均在 API 计时区间之外。",
                Error = error
            };
        }
        finally
        {
            if (Directory.Exists(tempDirectory) && CountFiles(tempDirectory) == 0)
                Directory.Delete(tempDirectory, true);
        }

        var json = JsonSerializer.Serialize(report, new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });
        File.WriteAllText(reportPath, json, new UTF8Encoding(false));
        Console.WriteLine(json);
        return status == "success" ? 0 : 1;
    }

    private static bool IsOneOf(string value, string first, string second)
        => string.Equals(value, first, StringComparison.OrdinalIgnoreCase)
           || string.Equals(value, second, StringComparison.OrdinalIgnoreCase);

    private static IEnumerable<StreamingProbeRow> CreateRows(string kind, int rows, ScenarioObservation observation)
    {
        for (var row = 0; row < rows; row++)
        {
            if (row == 0)
                observation.MarkFirstRow();
            yield return CreateRow(kind, row);
        }
    }

    private static StreamingProbeRow CreateRow(string kind, int row) => new()
    {
        C0 = GetValue(kind, row, 0), C1 = GetValue(kind, row, 1),
        C2 = GetValue(kind, row, 2), C3 = GetValue(kind, row, 3),
        C4 = GetValue(kind, row, 4), C5 = GetValue(kind, row, 5),
        C6 = GetValue(kind, row, 6), C7 = GetValue(kind, row, 7),
        C8 = GetValue(kind, row, 8), C9 = GetValue(kind, row, 9)
    };

    private static string GetValue(string kind, int row, int column)
        => kind == "repeated"
            ? $"重复-c{column + 1:D2}-中文"
            : $"r{row + 1:D6}-c{column + 1:D2}-中文";

    private static void ValidateRow(StreamingProbeRow item, string kind, int row)
    {
        for (var column = 0; column < ColumnCount; column++)
        {
            var actual = GetColumn(item, column);
            var expected = GetValue(kind, row, column);
            if (!string.Equals(actual, expected, StringComparison.Ordinal))
                throw new InvalidDataException($"单元格值不符: row={row}, column={column}, expected={expected}, actual={actual}");
        }
    }

    private static string GetColumn(StreamingProbeRow row, int index) => index switch
    {
        0 => row.C0, 1 => row.C1, 2 => row.C2, 3 => row.C3, 4 => row.C4,
        5 => row.C5, 6 => row.C6, 7 => row.C7, 8 => row.C8, 9 => row.C9,
        _ => throw new ArgumentOutOfRangeException(nameof(index))
    };

    private static void ValidateExportedWorkbook(string path, string kind, int rows)
    {
        using var input = File.OpenRead(path);
        using var workbook = new XSSFWorkbook(input);
        var sheet = workbook.GetSheet("Data") ?? throw new InvalidDataException("导出缺少 Data 工作表。");
        if (sheet.LastRowNum != rows)
            throw new InvalidDataException($"导出行数不符: expected={rows}, actual={sheet.LastRowNum}");

        var header = sheet.GetRow(0) ?? throw new InvalidDataException("导出缺少表头。");
        var columns = new int[ColumnCount];
        for (var column = 0; column < ColumnCount; column++)
        {
            columns[column] = FindColumn(header, "C" + column.ToString(CultureInfo.InvariantCulture));
            if (columns[column] < 0)
                throw new InvalidDataException($"导出缺少列 C{column}。");
        }

        for (var rowIndex = 0; rowIndex < rows; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex + 1) ?? throw new InvalidDataException($"导出缺少第 {rowIndex + 1} 行。");
            for (var column = 0; column < ColumnCount; column++)
            {
                var cell = row.GetCell(columns[column]);
                var actual = cell?.CellType == CellType.String ? cell.StringCellValue : null;
                var expected = GetValue(kind, rowIndex, column);
                if (!string.Equals(actual, expected, StringComparison.Ordinal))
                    throw new InvalidDataException($"导出单元格值不符: row={rowIndex}, column={column}");
            }
        }
    }

    private static int FindColumn(IRow header, string name)
    {
        for (var index = 0; index < header.LastCellNum; index++)
            if (string.Equals(header.GetCell(index)?.StringCellValue, name, StringComparison.Ordinal))
                return index;
        return -1;
    }

    private static void CreateSharedStringFixture(string path, string kind, int rows)
    {
        using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None,
            65536, FileOptions.SequentialScan);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false);
        WriteRootRelationships(archive);
        WriteContentTypes(archive);
        WriteWorkbook(archive);
        WriteWorkbookRelationships(archive);
        WriteStyles(archive);
        WriteSharedStrings(archive, kind, rows);
        WriteWorksheet(archive, kind, rows);
    }

    private static void WriteRootRelationships(ZipArchive archive) => WriteXml(archive, "_rels/.rels", writer =>
    {
        writer.WriteStartElement("Relationships", PackageRelationshipNamespace);
        WriteRelationship(writer, "rId1", DocumentRelationshipNamespace + "/officeDocument", "xl/workbook.xml");
        writer.WriteEndElement();
    });

    private static void WriteContentTypes(ZipArchive archive) => WriteXml(archive, "[Content_Types].xml", writer =>
    {
        writer.WriteStartElement("Types", ContentTypeNamespace);
        writer.WriteStartElement("Default", ContentTypeNamespace);
        writer.WriteAttributeString("Extension", "rels");
        writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
        writer.WriteEndElement();
        writer.WriteStartElement("Default", ContentTypeNamespace);
        writer.WriteAttributeString("Extension", "xml");
        writer.WriteAttributeString("ContentType", "application/xml");
        writer.WriteEndElement();
        WriteOverride(writer, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        WriteOverride(writer, "/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        WriteOverride(writer, "/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
        WriteOverride(writer, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        writer.WriteEndElement();
    });

    private static void WriteWorkbook(ZipArchive archive) => WriteXml(archive, "xl/workbook.xml", writer =>
    {
        writer.WriteStartElement("workbook", SpreadsheetNamespace);
        writer.WriteAttributeString("xmlns", "r", null, DocumentRelationshipNamespace);
        writer.WriteStartElement("sheets", SpreadsheetNamespace);
        writer.WriteStartElement("sheet", SpreadsheetNamespace);
        writer.WriteAttributeString("name", "Data");
        writer.WriteAttributeString("sheetId", "1");
        writer.WriteAttributeString("r", "id", DocumentRelationshipNamespace, "rId1");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
    });

    private static void WriteWorkbookRelationships(ZipArchive archive) => WriteXml(archive, "xl/_rels/workbook.xml.rels", writer =>
    {
        writer.WriteStartElement("Relationships", PackageRelationshipNamespace);
        WriteRelationship(writer, "rId1", DocumentRelationshipNamespace + "/worksheet", "worksheets/sheet1.xml");
        WriteRelationship(writer, "rId2", DocumentRelationshipNamespace + "/sharedStrings", "sharedStrings.xml");
        WriteRelationship(writer, "rId3", DocumentRelationshipNamespace + "/styles", "styles.xml");
        writer.WriteEndElement();
    });

    private static void WriteStyles(ZipArchive archive) => WriteXml(archive, "xl/styles.xml", writer =>
    {
        writer.WriteStartElement("styleSheet", SpreadsheetNamespace);
        writer.WriteStartElement("cellXfs", SpreadsheetNamespace);
        writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("xf", SpreadsheetNamespace);
        writer.WriteAttributeString("numFmtId", "0");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
    });

    private static void WriteSharedStrings(ZipArchive archive, string kind, int rows) =>
        WriteXml(archive, "xl/sharedStrings.xml", writer =>
        {
            var uniqueCount = kind == "repeated" ? ColumnCount : checked(rows * ColumnCount);
            writer.WriteStartElement("sst", SpreadsheetNamespace);
            writer.WriteAttributeString("count", checked(rows * ColumnCount).ToString(CultureInfo.InvariantCulture));
            writer.WriteAttributeString("uniqueCount", uniqueCount.ToString(CultureInfo.InvariantCulture));
            if (kind == "repeated")
            {
                for (var column = 0; column < ColumnCount; column++)
                    WriteSharedString(writer, GetValue(kind, 0, column));
            }
            else
            {
                for (var row = 0; row < rows; row++)
                    for (var column = 0; column < ColumnCount; column++)
                        WriteSharedString(writer, GetValue(kind, row, column));
            }
            writer.WriteEndElement();
        });

    private static void WriteWorksheet(ZipArchive archive, string kind, int rows) =>
        WriteXml(archive, "xl/worksheets/sheet1.xml", writer =>
        {
            writer.WriteStartElement("worksheet", SpreadsheetNamespace);
            writer.WriteStartElement("sheetData", SpreadsheetNamespace);
            writer.WriteStartElement("row", SpreadsheetNamespace);
            writer.WriteAttributeString("r", "1");
            for (var column = 0; column < ColumnCount; column++)
            {
                writer.WriteStartElement("c", SpreadsheetNamespace);
                writer.WriteAttributeString("r", CellReference(column, 1));
                writer.WriteAttributeString("t", "inlineStr");
                writer.WriteStartElement("is", SpreadsheetNamespace);
                writer.WriteElementString("t", SpreadsheetNamespace, "C" + column.ToString(CultureInfo.InvariantCulture));
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            writer.WriteEndElement();

            for (var row = 0; row < rows; row++)
            {
                var sheetRow = row + 2;
                writer.WriteStartElement("row", SpreadsheetNamespace);
                writer.WriteAttributeString("r", sheetRow.ToString(CultureInfo.InvariantCulture));
                for (var column = 0; column < ColumnCount; column++)
                {
                    var sharedIndex = kind == "repeated" ? column : checked(row * ColumnCount + column);
                    writer.WriteStartElement("c", SpreadsheetNamespace);
                    writer.WriteAttributeString("r", CellReference(column, sheetRow));
                    writer.WriteAttributeString("t", "s");
                    writer.WriteElementString("v", SpreadsheetNamespace, sharedIndex.ToString(CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            writer.WriteEndElement();
        });

    private static void WriteXml(ZipArchive archive, string path, Action<XmlWriter> write)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            CloseOutput = true,
            Indent = false,
            OmitXmlDeclaration = false
        });
        writer.WriteStartDocument();
        write(writer);
        writer.WriteEndDocument();
    }

    private static void WriteRelationship(XmlWriter writer, string id, string type, string target)
    {
        writer.WriteStartElement("Relationship", PackageRelationshipNamespace);
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Type", type);
        writer.WriteAttributeString("Target", target);
        writer.WriteEndElement();
    }

    private static void WriteOverride(XmlWriter writer, string name, string contentType)
    {
        writer.WriteStartElement("Override", ContentTypeNamespace);
        writer.WriteAttributeString("PartName", name);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static void WriteSharedString(XmlWriter writer, string value)
    {
        writer.WriteStartElement("si", SpreadsheetNamespace);
        writer.WriteElementString("t", SpreadsheetNamespace, value);
        writer.WriteEndElement();
    }

    private static string CellReference(int column, int row)
        => ((char)('A' + column)).ToString() + row.ToString(CultureInfo.InvariantCulture);

    private static bool ObserveSharedStringSpill(string tempDirectory, int rows)
    {
        var expectedIndexBytes = (long)rows * ColumnCount * 12;
        try
        {
            foreach (var session in Directory.GetDirectories(tempDirectory, "bing-stream-*"))
            {
                var files = Directory.GetFiles(session, "*.tmp");
                if (files.Length < 3)
                    continue;
                var inputExists = files.Any(file => Path.GetFileNameWithoutExtension(file) == "0"
                                                    && new FileInfo(file).Length > 0);
                var dataExists = files.Any(file => Path.GetFileNameWithoutExtension(file) == "1"
                                                   && new FileInfo(file).Length > 0);
                var indexExists = files.Any(file => Path.GetFileNameWithoutExtension(file) == "2"
                                                    && new FileInfo(file).Length == expectedIndexBytes);
                if (inputExists && dataExists && indexExists)
                    return true;
            }
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
        return false;
    }

    private static IEnumerable<string> GetFiles(string directory)
    {
        try
        {
            return Directory.Exists(directory)
                ? Directory.EnumerateFiles(directory, "*", SearchOption.AllDirectories).ToArray()
                : Array.Empty<string>();
        }
        catch (IOException) { return Array.Empty<string>(); }
        catch (UnauthorizedAccessException) { return Array.Empty<string>(); }
    }

    private static long GetDirectoryBytes(string directory, ISet<string> excludedFiles = null)
    {
        long bytes = 0;
        foreach (var path in GetFiles(directory))
        {
            if (excludedFiles != null && excludedFiles.Contains(path))
                continue;
            try { bytes = checked(bytes + new FileInfo(path).Length); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
        return bytes;
    }

    private static int CountFiles(string directory) => GetFiles(directory).Count();

    private static int CountNewFiles(string directory, ISet<string> baseline)
        => GetFiles(directory).Count(path => !baseline.Contains(path));

    private sealed class ScenarioObservation
    {
        private long _firstRowTimestamp;
        internal long FirstRowTimestamp => Interlocked.Read(ref _firstRowTimestamp);
        internal void MarkFirstRow() => Interlocked.CompareExchange(ref _firstRowTimestamp, Stopwatch.GetTimestamp(), 0);
    }

    internal sealed class CommitTemporaryFileTracker
    {
        private readonly string _directory;
        private readonly string _prefix;
        private readonly HashSet<string> _baseline;

        internal CommitTemporaryFileTracker(string targetPath)
        {
            var fullPath = Path.GetFullPath(targetPath);
            _directory = Path.GetDirectoryName(fullPath);
            _prefix = Path.GetFileName(fullPath) + ".";
            _baseline = new HashSet<string>(EnumerateMatchingFiles(), StringComparer.OrdinalIgnoreCase);
        }

        internal int GetCurrentCount() => GetCurrentFiles().Length;

        internal long GetCurrentBytes(string activePath = null, long activeStreamLength = 0)
        {
            long bytes = 0;
            foreach (var path in GetCurrentFiles())
            {
                long fileBytes = 0;
                try { fileBytes = new FileInfo(path).Length; }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
                if (!string.IsNullOrEmpty(activePath)
                    && string.Equals(Path.GetFullPath(path), Path.GetFullPath(activePath),
                        StringComparison.OrdinalIgnoreCase))
                    fileBytes = Math.Max(fileBytes, Math.Max(0, activeStreamLength));
                bytes = checked(bytes + fileBytes);
            }
            return bytes;
        }

        private string[] GetCurrentFiles()
            => EnumerateMatchingFiles().Where(path => !_baseline.Contains(path)).ToArray();

        private IEnumerable<string> EnumerateMatchingFiles()
        {
            if (!Directory.Exists(_directory))
                return Array.Empty<string>();

            return Directory.EnumerateFiles(_directory).Where(path =>
            {
                var name = Path.GetFileName(path);
                if (!name.StartsWith(_prefix, StringComparison.OrdinalIgnoreCase)
                    || !name.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase)
                    || name.Length != _prefix.Length + 32 + ".tmp".Length)
                    return false;

                var guid = name.Substring(_prefix.Length, name.Length - _prefix.Length - 4);
                return guid.Length == 32 && guid.All(Uri.IsHexDigit);
            }).ToArray();
        }
    }

    internal sealed class SamplingFileExportCommitter : IFileExportCommitter
    {
        private readonly IFileExportCommitter _inner;
        private readonly Action<string, long> _sampleAfterWrite;

        internal SamplingFileExportCommitter(IFileExportCommitter inner, Action<string, long> sampleAfterWrite)
        {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            _sampleAfterWrite = sampleAfterWrite ?? throw new ArgumentNullException(nameof(sampleAfterWrite));
        }

        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
            => _inner.Commit(path, stream =>
            {
                using var sampledStream = new SamplingWriteStream(stream, _sampleAfterWrite);
                write(sampledStream);
            }, cancellationToken, format);

        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format)
            => _inner.CommitAsync(path, async (stream, token) =>
            {
                using var sampledStream = new SamplingWriteStream(stream, _sampleAfterWrite);
                await writeAsync(sampledStream, token).ConfigureAwait(false);
            }, cancellationToken, format);
    }

    private sealed class SamplingWriteStream : Stream
    {
        private readonly Stream _inner;
        private readonly Action<string, long> _sampleAfterWrite;

        internal SamplingWriteStream(Stream inner, Action<string, long> sampleAfterWrite)
        {
            _inner = inner;
            _sampleAfterWrite = sampleAfterWrite;
        }

        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => _inner.CanSeek;
        public override bool CanWrite => _inner.CanWrite;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }

        public override void Flush() => _inner.Flush();
        public override Task FlushAsync(CancellationToken cancellationToken)
            => _inner.FlushAsync(cancellationToken);
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override int Read(Span<byte> buffer) => _inner.Read(buffer);
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
            => _inner.ReadAsync(buffer, cancellationToken);
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => _inner.SetLength(value);

        public override void Write(byte[] buffer, int offset, int count)
        {
            _inner.Write(buffer, offset, count);
            SampleAfterWrite();
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            _inner.Write(buffer);
            SampleAfterWrite();
        }

        public override void WriteByte(byte value)
        {
            _inner.WriteByte(value);
            SampleAfterWrite();
        }

        public override async Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            await _inner.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
            SampleAfterWrite();
        }

        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            await _inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
            SampleAfterWrite();
        }

        private void SampleAfterWrite()
        {
            var activePath = _inner is FileStream fileStream ? fileStream.Name : null;
            var activeLength = _inner.CanSeek ? _inner.Length : 0;
            _sampleAfterWrite(activePath, activeLength);
        }
    }

    internal sealed class ResourceSampler : IDisposable
    {
        private readonly string _libraryTempDirectory;
        private readonly string _npoiTempDirectory;
        private readonly HashSet<string> _npoiBaseline;
        private readonly CommitTemporaryFileTracker _commitTemporaryFiles;
        private readonly Process _process = Process.GetCurrentProcess();
        private readonly object _gate = new();
        private Timer _timer;
        private long _sampledPeakWorkingSet;
        private long _peakManagedBytes;
        private long _libraryTempPeak;
        private long _librarySessionTempPeak;
        private long _commitTempPeak;
        private long _npoiTempPeak;
        private bool _sessionAndCommitOverlapObserved;

        internal ResourceSampler(string libraryTempDirectory, string npoiTempDirectory, HashSet<string> npoiBaseline,
            CommitTemporaryFileTracker commitTemporaryFiles)
        {
            _libraryTempDirectory = libraryTempDirectory;
            _npoiTempDirectory = npoiTempDirectory;
            _npoiBaseline = npoiBaseline;
            _commitTemporaryFiles = commitTemporaryFiles;
        }

        internal long ProcessPeakWorkingSet { get; private set; }

        internal void Start()
        {
            Capture();
            _timer = new Timer(_ => Capture(), null, TimeSpan.FromMilliseconds(20), TimeSpan.FromMilliseconds(20));
        }

        internal ResourceSnapshot Stop()
        {
            _timer?.Dispose();
            _timer = null;
            Capture();
            lock (_gate)
            {
                _process.Refresh();
                ProcessPeakWorkingSet = _process.PeakWorkingSet64;
                return new ResourceSnapshot(_sampledPeakWorkingSet, _peakManagedBytes, _libraryTempPeak,
                    _librarySessionTempPeak, _commitTempPeak, _npoiTempPeak,
                    _sessionAndCommitOverlapObserved);
            }
        }

        internal void SampleAfterCommitWrite(string activePath, long activeStreamLength)
        {
            try
            {
                lock (_gate)
                    CaptureLibraryTemporaryFiles(activePath, activeStreamLength);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private void Capture()
        {
            try
            {
                lock (_gate)
                {
                    _process.Refresh();
                    _sampledPeakWorkingSet = Math.Max(_sampledPeakWorkingSet, _process.WorkingSet64);
                    _peakManagedBytes = Math.Max(_peakManagedBytes, GC.GetTotalMemory(false));
                    CaptureLibraryTemporaryFiles();
                    _npoiTempPeak = Math.Max(_npoiTempPeak,
                        GetDirectoryBytes(_npoiTempDirectory, _npoiBaseline));
                }
            }
            catch (InvalidOperationException) { }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private void CaptureLibraryTemporaryFiles(string activePath = null, long activeStreamLength = 0)
        {
            var sessionBytes = GetDirectoryBytes(_libraryTempDirectory);
            var commitBytes = _commitTemporaryFiles.GetCurrentBytes(activePath, activeStreamLength);
            _librarySessionTempPeak = Math.Max(_librarySessionTempPeak, sessionBytes);
            _commitTempPeak = Math.Max(_commitTempPeak, commitBytes);
            _libraryTempPeak = Math.Max(_libraryTempPeak, checked(sessionBytes + commitBytes));
            if (sessionBytes > 0 && commitBytes > 0)
                _sessionAndCommitOverlapObserved = true;
        }

        public void Dispose()
        {
            _timer?.Dispose();
            _process.Dispose();
        }
    }

    internal readonly struct ResourceSnapshot
    {
        internal ResourceSnapshot(long sampledPeakWorkingSet, long peakManagedBytes, long libraryTempPeak,
            long librarySessionTempPeak, long commitTempPeak, long npoiTempPeak,
            bool sessionAndCommitOverlapObserved)
        {
            SampledPeakWorkingSet = sampledPeakWorkingSet;
            PeakManagedBytes = peakManagedBytes;
            LibraryTempPeak = libraryTempPeak;
            LibrarySessionTempPeak = librarySessionTempPeak;
            CommitTempPeak = commitTempPeak;
            NpoiTempPeak = npoiTempPeak;
            SessionAndCommitOverlapObserved = sessionAndCommitOverlapObserved;
        }

        internal long SampledPeakWorkingSet { get; }
        internal long PeakManagedBytes { get; }
        internal long LibraryTempPeak { get; }
        internal long LibrarySessionTempPeak { get; }
        internal long CommitTempPeak { get; }
        internal long NpoiTempPeak { get; }
        internal bool SessionAndCommitOverlapObserved { get; }
    }

    private sealed class ExcelStreamingScenarioReport
    {
        public int Rows { get; init; }
        public int Columns { get; init; }
        public string Kind { get; init; }
        public string Direction { get; init; }
        public double ElapsedMs { get; init; }
        public double RowsPerSecond { get; init; }
        public double? FirstRowMs { get; init; }
        public long PeakWorkingSet { get; init; }
        public long SampledPeakWorkingSet { get; init; }
        public long PeakManagedBytes { get; init; }
        public long LibraryTempPeak { get; init; }
        public long LibrarySessionTempPeak { get; init; }
        public long CommitTempPeak { get; init; }
        public long NpoiTempPeak { get; init; }
        public bool CommitOverlapObserved { get; init; }
        public bool SpillObserved { get; init; }
        public int TempFilesRemaining { get; init; }
        public int LibraryTempFilesRemaining { get; init; }
        public int LibrarySessionTempFilesRemaining { get; init; }
        public int CommitTempFilesRemaining { get; init; }
        public int NpoiTempFilesRemaining { get; init; }
        public string Status { get; init; }
        public string Runtime { get; init; }
        public string OperatingSystem { get; init; }
        public string OSArchitecture { get; init; }
        public string ProcessArchitecture { get; init; }
        public int ProcessorCount { get; init; }
        public bool Is64BitProcess { get; init; }
        public string MeasurementNotes { get; init; }
        public string Error { get; init; }
    }

    private sealed class StreamingProbeRow
    {
        public string C0 { get; set; }
        public string C1 { get; set; }
        public string C2 { get; set; }
        public string C3 { get; set; }
        public string C4 { get; set; }
        public string C5 { get; set; }
        public string C6 { get; set; }
        public string C7 { get; set; }
        public string C8 { get; set; }
        public string C9 { get; set; }
    }
}
