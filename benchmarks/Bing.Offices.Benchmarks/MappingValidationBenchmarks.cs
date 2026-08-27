using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Bing.Offices.Configurations;
using Bing.Offices.Attributes;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Bing.Offices.Extensions;
using Microsoft.Extensions.DependencyInjection;
using System.Threading;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// Mapping、Unique、配置解析和 Profile 注册基准矩阵。
/// </summary>
[MemoryDiagnoser]
[SimpleJob(launchCount: 1, warmupCount: 2, iterationCount: 3)]
public class MappingValidationBenchmarks
{
    private const long PeakWorkingSetCeilingBytes = 1024L * 1024 * 1024;
    private static int _workingSetEvidenceWritten;
    private ExcelMappingPlanFactory _planFactory = null!;
    private ExcelMappingConfiguration _configuration = null!;
    private ExcelMappingDocument _document = null!;
    private string _json = string.Empty;
    private string _xml = string.Empty;
    private string _jsonV1 = string.Empty;
    private string _xmlV1 = string.Empty;
    private ExcelMappingConfiguration _multiRuleConfiguration = null!;
    private ExcelMappingConfiguration _profileConfiguration = null!;
    private JsonSerializerOptions _cacheKeySerializerOptions = null!;

    /// <summary>
    /// 初始化矩阵输入。
    /// </summary>
    [GlobalSetup]
    public void Setup()
    {
        var namedRules = Enumerable.Range(0, 10000)
            .Select(index => (INamedExcelValidationRule)new BenchmarkNamedValidationRule($"rule-{index}"))
            .ToArray();
        _planFactory = new ExcelMappingPlanFactory(namedValidationRules: namedRules);
        _cacheKeySerializerOptions = new JsonSerializerOptions { IgnoreNullValues = false };
        _configuration = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration { PropertyName = nameof(BenchmarkRow.Code), Title = "编码" },
                new ExcelColumnConfiguration { PropertyName = nameof(BenchmarkRow.Amount), Title = "金额" }
            }
        };
        var profileBuilder = new ImportMappingBuilder<BenchmarkRow>();
        profileBuilder.Property(row => row.Code).HasHeader("编码");
        _profileConfiguration = profileBuilder.Build();
        _document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration
            {
                Profile = "benchmarks",
                ModelAlias = "benchmark-row"
            },
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

    /// <summary>测量 JSON v2 配置解析。</summary>
    [Benchmark]
    public int ParseJsonV2() => ExcelMappingConfigurationLoader.FromJsonDocument(_json).Import.Columns.Count;

    /// <summary>测量 XML v2 配置解析。</summary>
    [Benchmark]
    public int ParseXmlV2() => ExcelMappingConfigurationLoader.FromXmlDocument(_xml).Import.Columns.Count;

    /// <summary>测量 JSON v1 迁移解析。</summary>
    [Benchmark]
    public int ParseJsonV1() => ExcelMappingConfigurationLoader.MigrateV1Json(_jsonV1,
        MappingDirection.Import).Import.Columns.Count;

    /// <summary>测量 XML v1 迁移解析。</summary>
    [Benchmark]
    public int ParseXmlV1() => ExcelMappingConfigurationLoader.MigrateV1Xml(_xmlV1,
        MappingDirection.Import).Import.Columns.Count;

    /// <summary>测量 10K 命名规则配置的不可变计划构建。</summary>
    [Benchmark]
    public int MultiRulePlanBuild() => _planFactory.Create<BenchmarkRow>(
        new ExcelMappingDocument
        {
            Import = _multiRuleConfiguration
        }, MappingDirection.Import).Columns.Count;

    /// <summary>测量旧式 JSON 字符串转 UTF-8 的缓存键序列化路径。</summary>
    [Benchmark]
    public string CacheKeyStringToUtf8()
    {
        var json = JsonSerializer.Serialize(CreateCacheKeyPayload(), _cacheKeySerializerOptions);
        var payload = Encoding.UTF8.GetBytes(json);
        return Convert.ToBase64String(SHA256.HashData(payload));
    }

    /// <summary>测量直接序列化 UTF-8 字节的缓存键路径。</summary>
    [Benchmark]
    public string CacheKeyUtf8Bytes()
    {
        var payload = JsonSerializer.SerializeToUtf8Bytes(CreateCacheKeyPayload(), _cacheKeySerializerOptions);
        return Convert.ToBase64String(SHA256.HashData(payload));
    }

    /// <summary>读取当前进程已观测峰值工作集，供资源证据采集。</summary>
    [Benchmark]
    public long PeakWorkingSetBytes()
    {
        var value = Process.GetCurrentProcess().PeakWorkingSet64;
        EnsureResourceCeiling("PeakWorkingSetBytes", value, PeakWorkingSetCeilingBytes);
        return value;
    }

    private static void EnsureResourceCeiling(string metric, long value, long ceiling)
    {
        if (value > ceiling)
            throw new InvalidOperationException($"RESOURCE_CEILING failed metric={metric} value={value} ceiling={ceiling}");
        var shouldWrite = Interlocked.Exchange(ref _workingSetEvidenceWritten, 1) == 0;
        if (shouldWrite)
            Console.WriteLine($"RESOURCE_METRIC metric={metric} value={value} ceiling={ceiling} status=passed");
    }

    /// <summary>测量显式 Profile 注册解析。</summary>
    [Benchmark]
    public int ExplicitRegistration()
    {
        var registry = new MappingProfileRegistry();
        registry.Register(new ProfileDescriptor("benchmarks", MappingDirection.Import,
            typeof(BenchmarkRow), _profileConfiguration));
        return registry.TryGetDescriptor("benchmarks", MappingDirection.Import, typeof(BenchmarkRow), out _)
            ? 1 : 0;
    }

    /// <summary>测量程序集扫描注册入口。</summary>
    [Benchmark]
    public int AssemblyScanRegistration()
    {
        var services = new ServiceCollection();
        services.AddMappingProfiles(typeof(MappingValidationBenchmarks).Assembly);
        return services.Count;
    }

    private object CreateCacheKeyPayload() => new
    {
        _document.TenantId,
        ModelType = typeof(BenchmarkRow).AssemblyQualifiedName,
        Direction = MappingDirection.Import,
        _document.ConfigurationVersion,
        Configuration = _configuration
    };

    private sealed class BenchmarkProfile : IImportMappingProfile<BenchmarkRow>
    {
        public void Configure(ImportMappingBuilder<BenchmarkRow> setting)
        {
            setting.Property(row => row.Code).HasHeader("编码");
        }
    }

    private sealed class BenchmarkNamedValidationRule : INamedExcelValidationRule
    {
        public BenchmarkNamedValidationRule(string name) => Name = name;

        public string Name { get; }
        public string ErrorMessage => "benchmark";
        public bool Validate(ExcelValidationContext context) => true;
    }

    private sealed class BenchmarkRow
    {
        public string Code { get; set; } = string.Empty;
        public decimal Amount { get; set; }
    }
}

/// <summary>
/// 动态映射计划构建与缓存基准。
/// </summary>
[MemoryDiagnoser]
[SimpleJob(launchCount: 1, warmupCount: 2, iterationCount: 3)]
public class DynamicPlanBenchmarks
{
    [Params(100, 500)]
    public int PlanBuildCount { get; set; }

    private ExcelMappingPlanFactory _planFactory = null!;
    private ExcelMappingConfiguration _profileConfiguration = null!;

    [GlobalSetup]
    public void Setup()
    {
        _planFactory = new ExcelMappingPlanFactory(namedValidationRules:
            Enumerable.Range(0, 10000)
                .Select(index => (INamedExcelValidationRule)new BenchmarkNamedValidationRule($"rule-{index}"))
                .ToArray());
        var profileBuilder = new ImportMappingBuilder<BenchmarkRow>();
        profileBuilder.Property(row => row.Code).HasHeader("编码");
        _profileConfiguration = profileBuilder.Build();
    }

    [Benchmark]
    public int DynamicPlanBuildCold()
    {
        var count = 0;
        for (var index = 0; index < PlanBuildCount; index++)
        {
            var factory = CreatePlanFactory();
            count += factory.Create<BenchmarkRow>(CreateDynamicDocument(index), MappingDirection.Import).Columns.Count;
        }
        return count;
    }

    [Benchmark]
    public int DynamicPlanBuildCacheHit()
    {
        var count = _planFactory.Create<BenchmarkRow>(CreateDynamicDocument(0), MappingDirection.Import).Columns.Count;
        for (var index = 1; index < PlanBuildCount; index++)
            count += _planFactory.Create<BenchmarkRow>(CreateDynamicDocument(0), MappingDirection.Import).Columns.Count;
        return count;
    }

    [Benchmark]
    public int DynamicPlanBuildCacheMiss()
    {
        var count = 0;
        for (var index = 0; index < PlanBuildCount; index++)
            count += _planFactory.Create<BenchmarkRow>(CreateDynamicDocument(index), MappingDirection.Import).Columns.Count;
        return count;
    }

    private ExcelMappingPlanFactory CreatePlanFactory() => new(namedValidationRules:
        Enumerable.Range(0, 10000)
            .Select(index => (INamedExcelValidationRule)new BenchmarkNamedValidationRule($"rule-{index}"))
            .ToArray());

    private ExcelMappingDocument CreateDynamicDocument(int index) => new()
    {
        TenantId = $"tenant-{index}",
        Import = new ExcelMappingConfiguration
        {
            Profile = "benchmarks",
            Columns = _profileConfiguration.Columns
        }
    };

    private sealed class BenchmarkNamedValidationRule : INamedExcelValidationRule
    {
        public BenchmarkNamedValidationRule(string name) => Name = name;
        public string Name { get; }
        public string ErrorMessage => "benchmark";
        public bool Validate(ExcelValidationContext context) => true;
    }

    private sealed class BenchmarkRow
    {
        public string Code { get; set; } = string.Empty;
    }
}

/// <summary>
/// 租户映射计划缓存隔离与淘汰基准。
/// </summary>
[MemoryDiagnoser]
[SimpleJob(launchCount: 1, warmupCount: 2, iterationCount: 3)]
public class TenantPlanCacheBenchmarks
{
    [Params(100, 1000)]
    public int TenantCount { get; set; }

    private ExcelMappingDocument _document = null!;

    [GlobalSetup]
    public void Setup()
    {
        _document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration { Profile = "benchmarks", ModelAlias = "benchmark-row" },
            Export = new ExcelMappingConfiguration()
        };
    }

    [Benchmark]
    public int TenantPlanCache()
    {
        var factory = new ExcelMappingPlanFactory(cacheCapacity: TenantCount);
        var count = 0;
        for (var index = 0; index < TenantCount; index++)
            count += factory.Create<BenchmarkRow>(CreateTenantDocument(index), MappingDirection.Import).Columns.Count;
        return count;
    }

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
        return count + rebuilt.Columns.Count + evictedAndRebuilt;
    }

    private ExcelMappingDocument CreateTenantDocument(int index) => new()
    {
        Version = _document.Version,
        TenantId = $"tenant-{index}",
        Import = new ExcelMappingConfiguration
        {
            Profile = _document.Import.Profile,
            ModelAlias = _document.Import.ModelAlias
        },
        Export = _document.Export
    };

    private sealed class BenchmarkRow
    {
        public string Code { get; set; } = string.Empty;
    }
}

/// <summary>
/// 有界 Regex 缓存命中基准。
/// </summary>
[MemoryDiagnoser]
[SimpleJob(launchCount: 1, warmupCount: 2, iterationCount: 3)]
public class RegexCacheBenchmarks
{
    private readonly ExcelValidationContext _context = new("CODE-123", "Sheet1", 2, 1, "Code");
    private readonly ExcelRegexAttribute _attribute = new("^CODE-[0-9]+$");
    private readonly RegexExcelValidationRule _rule = new();

    /// <summary>
    /// 测量重复命中有界 Regex 缓存的校验成本。
    /// </summary>
    [Benchmark]
    public bool RegexCacheHit() => _rule.Validate(_attribute, _context);
}

/// <summary>
/// Unique committed/pending journal 基准。
/// </summary>
[MemoryDiagnoser]
[SimpleJob(launchCount: 1, warmupCount: 2, iterationCount: 3)]
public class UniqueJournalBenchmarks
{
    /// <summary>
    /// 唯一列数量。
    /// </summary>
    [Params(1, 5)]
    public int UniqueColumnCount { get; set; }

    /// <summary>
    /// 唯一值行数。
    /// </summary>
    [Params(10000, 100000)]
    public int UniqueRowCount { get; set; }

    /// <summary>
    /// 测量 Unique committed/pending journal 的线性插入。
    /// </summary>
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
}
