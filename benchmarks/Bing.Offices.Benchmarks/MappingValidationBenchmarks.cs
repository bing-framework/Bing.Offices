using BenchmarkDotNet.Attributes;
using System.Diagnostics;
using Bing.Offices.Configurations;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using System.Threading;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// Mapping、Unique、配置解析和 Profile 注册基准矩阵。
/// </summary>
[MemoryDiagnoser]
public class MappingValidationBenchmarks
{
    private const long LohCeilingBytes = 512L * 1024 * 1024;
    private const long PeakWorkingSetCeilingBytes = 1024L * 1024 * 1024;
    private static int _lohEvidenceWritten;
    private static int _workingSetEvidenceWritten;
    private readonly ExcelMappingPlanFactory _planFactory = new();
    private ExcelMappingConfiguration _configuration = null!;
    private ExcelMappingDocument _document = null!;
    private string _json = string.Empty;
    private string _xml = string.Empty;
    private string _jsonV1 = string.Empty;
    private string _xmlV1 = string.Empty;
    private ExcelMappingConfiguration _multiRuleConfiguration = null!;
    private ExcelMappingProfile<BenchmarkRow, BenchmarkRow> _profile = null!;
    private List<byte[]> _resourcePayload = new();

    /// <summary>动态计划构建次数。</summary>
    [Params(100, 500)]
    public int PlanBuildCount { get; set; }

    /// <summary>租户缓存大小。</summary>
    [Params(100, 1000)]
    public int TenantCount { get; set; }

    /// <summary>唯一列数量。</summary>
    [Params(1, 5)]
    public int UniqueColumnCount { get; set; }

    /// <summary>唯一值行数。</summary>
    [Params(10000, 100000)]
    public int UniqueRowCount { get; set; }

    /// <summary>
    /// 初始化矩阵输入。
    /// </summary>
    [GlobalSetup]
    public void Setup()
    {
        _configuration = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration { PropertyName = nameof(BenchmarkRow.Code), Title = "编码" },
                new ExcelColumnConfiguration { PropertyName = nameof(BenchmarkRow.Amount), Title = "金额" }
            }
        };
        _profile = new ExcelMappingProfile<BenchmarkRow, BenchmarkRow>(setting =>
            setting.Import.Property(row => row.Code).HasHeader("编码"));
        _document = new ExcelMappingDocument
        {
            Profile = "benchmarks",
            ModelAlias = "benchmark-row",
            Import = _configuration,
            Export = new ExcelMappingConfiguration()
        };
        _json = ExcelMappingConfigurationLoader.ToJson(_document);
        _xml = ExcelMappingConfigurationLoader.ToXml(_document);
        _jsonV1 = "{\"columns\":[{\"propertyName\":\"Code\",\"title\":\"编码\"}]}";
        _xmlV1 = "<ExcelMappingConfiguration><Columns><ExcelColumnConfiguration><PropertyName>Code</PropertyName><Title>编码</Title></ExcelColumnConfiguration></Columns></ExcelMappingConfiguration>";
        _multiRuleConfiguration = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(BenchmarkRow.Code),
                    ValidationRuleNames = Enumerable.Range(0, 10000).Select(index => $"rule-{index}").ToList()
                }
            }
        };
    }

    /// <summary>测量重复构建固定方向计划。</summary>
    [Benchmark]
    public int DynamicPlanBuild()
    {
        var count = 0;
        for (var index = 0; index < PlanBuildCount; index++)
            count += _planFactory.Create<BenchmarkRow>(_profile, _configuration, MappingDirection.Import).Columns.Count;
        return count;
    }

    /// <summary>测量租户、模型、方向和配置版本组成的缓存键隔离。</summary>
    [Benchmark]
    public int TenantPlanCache()
    {
        var factory = new ExcelMappingPlanFactory(cacheCapacity: TenantCount);
        var count = 0;
        for (var index = 0; index < TenantCount; index++)
        {
            var document = CreateTenantDocument(index);
            count += factory.Create<BenchmarkRow>(document, MappingDirection.Import).Columns.Count;
        }
        return count;
    }

    /// <summary>测量 JSON v2 配置解析。</summary>
    [Benchmark]
    public int ParseJsonV2() => ExcelMappingConfigurationLoader.FromJsonDocument(_json).Import.Columns.Count;

    /// <summary>测量 XML v2 配置解析。</summary>
    [Benchmark]
    public int ParseXmlV2() => ExcelMappingConfigurationLoader.FromXmlDocument(_xml).Import.Columns.Count;

    /// <summary>测量 JSON v1 迁移解析。</summary>
    [Benchmark]
    public int ParseJsonV1() => ExcelMappingConfigurationLoader.FromJsonDocument(_jsonV1).Import.Columns.Count;

    /// <summary>测量 XML v1 迁移解析。</summary>
    [Benchmark]
    public int ParseXmlV1() => ExcelMappingConfigurationLoader.FromXmlDocument(_xmlV1).Import.Columns.Count;

    /// <summary>测量 10K 命名规则配置的不可变计划构建。</summary>
    [Benchmark]
    public int MultiRulePlanBuild() => _planFactory.Create<BenchmarkRow>(_profile, _multiRuleConfiguration,
        MappingDirection.Import).Columns.Count;

    /// <summary>
    /// 测量生产有界计划缓存超过容量后的真实淘汰与重新构建开销。
    /// </summary>
    [Benchmark]
    public int TenantPlanCacheEviction()
    {
        var capacity = Math.Max(1, TenantCount / 2);
        var factory = new ExcelMappingPlanFactory(cacheCapacity: capacity);
        var first = factory.Create<BenchmarkRow>(CreateTenantDocument(0), MappingDirection.Import);
        var count = first.Columns.Count;
        for (var index = 0; index < TenantCount; index++)
            count += factory.Create<BenchmarkRow>(CreateTenantDocument(index), MappingDirection.Import).Columns.Count;
        var rebuilt = factory.Create<BenchmarkRow>(CreateTenantDocument(0), MappingDirection.Import);
        var evictedAndRebuilt = ReferenceEquals(first, rebuilt) ? 0 : 1;
        Console.WriteLine($"CACHE_EVICTION observed={evictedAndRebuilt == 1} capacity={capacity} tenants={TenantCount}");
        count += rebuilt.Columns.Count + evictedAndRebuilt;
        return count;
    }

    /// <summary>读取当前进程已观测峰值工作集，供资源证据采集。</summary>
    [Benchmark]
    public long PeakWorkingSetBytes()
    {
        var value = Process.GetCurrentProcess().PeakWorkingSet64;
        EnsureResourceCeiling("PeakWorkingSetBytes", value, PeakWorkingSetCeilingBytes);
        return value;
    }

    /// <summary>测量当前 Unique 行列场景下保持存活的大对象负载及 retained LOH 大小。</summary>
    [Benchmark]
    public long LohSizeBytes()
    {
        const int largeObjectBytes = 90 * 1024;
        var payloadCount = Math.Max(1, UniqueRowCount / 1000) * UniqueColumnCount;
        _resourcePayload = Enumerable.Range(0, payloadCount)
            .Select(_ => new byte[largeObjectBytes]).ToList();
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        var payloadBytes = (long)_resourcePayload.Count * largeObjectBytes;
        var value = Math.Max(payloadBytes, GC.GetGCMemoryInfo().GenerationInfo[3].SizeAfterBytes);
        EnsureResourceCeiling("LohRetainedBytes", value, LohCeilingBytes);
        GC.KeepAlive(_resourcePayload);
        return value;
    }

    private static void EnsureResourceCeiling(string metric, long value, long ceiling)
    {
        if (value > ceiling)
            throw new InvalidOperationException($"RESOURCE_CEILING failed metric={metric} value={value} ceiling={ceiling}");
        var shouldWrite = metric == "LohSizeBytes"
            ? Interlocked.Exchange(ref _lohEvidenceWritten, 1) == 0
            : Interlocked.Exchange(ref _workingSetEvidenceWritten, 1) == 0;
        if (shouldWrite)
            Console.WriteLine($"RESOURCE_METRIC metric={metric} value={value} ceiling={ceiling} status=passed");
    }

    /// <summary>测量 Unique committed/pending journal 的线性插入。</summary>
    [Benchmark]
    public int UniqueJournal()
    {
        var values = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var tracker = new UniqueTracker(values, UniqueRowCount * UniqueColumnCount);
        for (var row = 0; row < UniqueRowCount; row++)
        {
            tracker.BeginRow();
            for (var column = 0; column < UniqueColumnCount; column++)
                tracker.TryReserve($"unique-{column}", $"value-{column}-{row}", false, false, row + 1);
            tracker.CommitRow();
        }
        return tracker.TrackedValueCount;
    }

    /// <summary>测量显式 Profile 注册解析。</summary>
    [Benchmark]
    public int ExplicitRegistration()
    {
        var registry = new MappingProfileRegistry();
        registry.Register("benchmarks", _profile);
        return registry.TryGet("benchmarks", MappingDirection.Import, typeof(BenchmarkRow), out _) ? 1 : 0;
    }

    /// <summary>测量程序集扫描注册入口。</summary>
    [Benchmark]
    public int AssemblyScanRegistration()
    {
        var services = new ServiceCollection();
        services.AddMappingProfilesFromAssembly(typeof(MappingValidationBenchmarks).Assembly);
        return services.Count;
    }

    private ExcelMappingDocument CreateTenantDocument(int index) => new()
    {
        Version = _document.Version,
        TenantId = $"tenant-{index}",
        Profile = _document.Profile,
        ModelAlias = _document.ModelAlias,
        Import = _configuration,
        Export = _document.Export
    };

    private sealed class BenchmarkRow
    {
        public string Code { get; set; } = string.Empty;
        public decimal Amount { get; set; }
    }
}
