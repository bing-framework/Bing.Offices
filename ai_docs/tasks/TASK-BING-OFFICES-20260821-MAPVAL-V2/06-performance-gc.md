# TASK-BING-OFFICES-20260821-MAPVAL-V2 性能与 GC 证据

## 状态

- Task ID：`TASK-BING-OFFICES-20260821-MAPVAL-V2`
- 当前状态：`PARTIAL`

## 已执行 Benchmark

命令：`dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-restore -- --filter "*StreamPipelineBenchmarks*" --job short`

环境：Windows 10 22H2，Intel Core Ultra 7 270K Plus，.NET 8.0.27，BenchmarkDotNet 0.14.0，X64 RyuJIT AVX2，Concurrent Workstation GC。

| 方法 | 行数 | 平均耗时 | Gen0 | Gen1 | Gen2 | 分配 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Import | 1000 | 15.90 ms | 984.3750 | 921.8750 | 93.7500 | 16.36 MB |
| Export | 1000 | 15.68 ms | 500.0000 | 343.7500 | 218.7500 | 8.26 MB |

结果文件：`BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.StreamPipelineBenchmarks-report-github.md`。

本轮结果明确存在 Gen0/Gen1/Gen2 回收，未宣称零 GC。LOH 和 Peak Working Set 的独立受控测量见 Round 7；此处历史 BenchmarkDotNet 报告本身不作为资源 ceiling 证据。

## Review Fix Round 1 新增矩阵

命令：
`dotnet "$env:TEMP\BingOfficesBenchmark\Bing.Offices.Benchmarks.dll" --filter "*MappingValidationBenchmarks.UniqueJournal*" --job short --warmupCount 1 --iterationCount 3`

环境：同上；ShortRun 为 3 次迭代、1 次预热、1 次启动。完整报告：`BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.MappingValidationBenchmarks-report-github.md`。

| PlanBuildCount | TenantCount | UniqueColumnCount | UniqueRowCount | Mean | Gen0 | Gen1 | Gen2 | Allocated |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 100 | 100 | 1 | 10,000 | 3.084 ms | 500.0000 | 414.0625 | 332.0313 | 7.03 MB |
| 100 | 100 | 1 | 100,000 | 48.610 ms | 4272.7273 | 1909.0909 | 1272.7273 | 68.75 MB |
| 100 | 100 | 5 | 10,000 | 23.376 ms | 2937.5000 | 1937.5000 | 1468.7500 | 33.03 MB |
| 100 | 100 | 5 | 100,000 | 291.667 ms | 18000.0000 | 7000.0000 | 3500.0000 | 322.39 MB |
| 500 | 1000 | 1 | 10,000 | 2.997 ms | 679.6875 | 351.5625 | 332.0313 | 7.03 MB |
| 500 | 1000 | 1 | 100,000 | 48.266 ms | 4545.4545 | 2272.7273 | 1272.7273 | 68.75 MB |
| 500 | 1000 | 5 | 10,000 | 23.278 ms | 2937.5000 | 1937.5000 | 1468.7500 | 33.03 MB |
| 500 | 1000 | 5 | 100,000 | 296.660 ms | 18000.0000 | 7000.0000 | 3500.0000 | 322.40 MB |

该基准的 `TenantCount` 和 `PlanBuildCount` 仅用于矩阵参数；`UniqueJournal` 本身按行和唯一列数执行，因此不能据此宣称租户缓存或 Plan Build 已完成性能验收。结果证明当前 journal 在 1 列 10K/100K 和 5 列 10K/100K 场景均能完成，且存在正常 GC；LOH 与 Peak Working Set 仍未测量。

## Review Fix Round 2 Mapping 矩阵

命令：
`dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --filter "*MappingValidationBenchmarks*" --job short --warmupCount 0 --iterationCount 1 --launchCount 1`

环境：Windows 10 22H2，Intel Core Ultra 7 270K Plus，.NET 8.0.27，BenchmarkDotNet 0.14.0，X64 RyuJIT AVX2，Concurrent Workstation GC。完整矩阵共 192 个执行组合，耗时约 609 秒。首次运行中 `MultiRulePlanBuild` 因 benchmark 本身传入 null Profile 失败，其余 176 个组合有结果；修复 benchmark 输入后单独重跑全部 16 个多规则组合，耗时约 54 秒并全部通过。原始日志：

- `BenchmarkDotNet.Artifacts/Bing.Offices.Benchmarks.MappingValidationBenchmarks-20260822-135406.log`
- `BenchmarkDotNet.Artifacts/Bing.Offices.Benchmarks.MappingValidationBenchmarks-20260822-140827.log`

以下是完整矩阵中 `PlanBuildCount=500、TenantCount=1000、UniqueColumnCount=5、UniqueRowCount=100000` 组合的代表结果；Ops/s 按 `1 / Mean` 换算，NA 表示 BenchmarkDotNet 未提供置信区间或该项不适用。

| 方法 | Mean | Ops/s | Gen0 | Gen1 | Gen2 | Allocated |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| DynamicPlanBuild | 410.107 us | 2,438 | 128.4180 | - | - | 2,420,000 B |
| TenantPlanCache | 942.753 us | 1,061 | 261.7188 | 97.6563 | - | 4,928,001 B |
| ParseJsonV2 | 6.408 us | 156,057 | 0.6714 | 0.0305 | - | 12,941 B |
| ParseXmlV2 | 15.089 us | 66,274 | 1.6174 | 0.0610 | - | 30,499 B |
| ParseJsonV1 | 1.324 us | 755,539 | 0.1659 | 0.0038 | 0.0019 | 3,174 B |
| ParseXmlV1 | 2.670 us | 374,596 | 1.4114 | 0.0610 | - | 26,665 B |
| MultiRulePlanBuild | 283.400 us | 3,529 | 124.5117 | 124.5117 | 124.5117 | 818.35 KB |
| TenantPlanCacheEviction | 977.726 us | 1,023 | 264.6484 | 205.0781 | - | 4,984,489 B |
| PeakWorkingSetBytes | 4.034 ms | 248 | - | - | - | 3,582 B |
| UniqueJournal | 291.301 ms | 3.43 | 17,000.0000 | 6,000.0000 | 3,000.0000 | 338,053,700 B |
| ExplicitRegistration | 60.69 ns | 16,476,600 | 0.0169 | - | - | 320 B |
| AssemblyScanRegistration | 403.01 ns | 2,481,300 | 0.0439 | - | - | 832 B |

### MultiRulePlanBuild 16 组

该方法使用 10,000 个 validation rule names；每组均成功，Mean 范围为 280.488–287.871 us，Gen0/Gen1/Gen2 均为 124.5117/124.5117/124.5117，Allocated 均为 818.35 KB。

| PlanBuildCount | TenantCount | UniqueColumnCount | UniqueRowCount | Mean |
| ---: | ---: | ---: | ---: | ---: |
| 100 | 100 | 1 | 10,000 | 283.526 us |
| 100 | 100 | 1 | 100,000 | 283.436 us |
| 100 | 100 | 5 | 10,000 | 280.794 us |
| 100 | 100 | 5 | 100,000 | 287.871 us |
| 100 | 1,000 | 1 | 10,000 | 284.420 us |
| 100 | 1,000 | 1 | 100,000 | 283.563 us |
| 100 | 1,000 | 5 | 10,000 | 280.488 us |
| 100 | 1,000 | 5 | 100,000 | 281.796 us |
| 500 | 100 | 1 | 10,000 | 285.147 us |
| 500 | 100 | 1 | 100,000 | 282.570 us |
| 500 | 100 | 5 | 10,000 | 283.341 us |
| 500 | 100 | 5 | 100,000 | 282.463 us |
| 500 | 1,000 | 1 | 10,000 | 286.071 us |
| 500 | 1,000 | 1 | 100,000 | 285.080 us |
| 500 | 1,000 | 5 | 10,000 | 283.959 us |
| 500 | 1,000 | 5 | 100,000 | 283.400 us |

`TenantPlanCache` 和 `TenantPlanCacheEviction` 是 benchmark 内部 Dictionary/Queue 场景，用于比较计划构建和淘汰操作开销，不代表生产 tenant cache 已实现或已完成容量验收。Round 2 的 `PeakWorkingSetBytes` 仍只是 BenchmarkDotNet 进程累计值，不能替代 Round 7 的独立进程测量。

## Review Fix Round 7 非 Dry 资源证据

命令：
`dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --resource-probe artifacts/mapval-v2/resource-probe-round7.jsonl`

该命令不是 BenchmarkDotNet Dry/占位运行。父进程按 `PlanBuildCount={100,500}`、`TenantCount={100,1000}`、`UniqueColumnCount={1,5}`、`UniqueRowCount={10000,100000}` 启动 16 个独立子进程；每个子进程真实构建 Plan、执行 Unique journal，并保持 `90 KiB` 大对象负载到测量点。每组记录 `payloadBytes`、GC LOH `SizeBefore/SizeAfter`、`lohPeakBytes`、子进程 `PeakWorkingSet64`、ceiling 和 exit code。

环境：Windows，X64，`.NET 8.0.27`，Release，进程路径为 `benchmarks/Bing.Offices.Benchmarks/bin/Release/net8.0/Bing.Offices.Benchmarks.exe`。ceiling 为 `LOH <= 512 MiB`、`Peak Working Set <= 1024 MiB`。

| PlanBuildCount | TenantCount | UniqueColumnCount | UniqueRowCount | LohPeakBytes | PeakWorkingSetBytes | 判定 |
| ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 100/500 | 100/1000 | 1 | 10,000 | 2,057,136 | 42,352,640 | 4/4 通过 |
| 100/500 | 100/1000 | 1 | 100,000 | 18,570,808 | 67,629,056 | 4/4 通过 |
| 100/500 | 100/1000 | 5 | 10,000 | 9,400,376 | 58,040,320 | 4/4 通过 |
| 100/500 | 100/1000 | 5 | 100,000 | 85,801,616 | 133,619,712 | 4/4 通过 |

表中值为对应 4 个 Plan/Tenant 组合的最大值；每个组合的独立原始 JSONL 记录均保存在 `artifacts/mapval-v2/resource-probe-round7.jsonl`。原始记录包含 16 个 `child` 行及其嵌套 `scenario` JSON，所有 `exitCode=0`、`status=passed`。本轮最大 LOH 约 `81.8 MiB`，最大 Peak Working Set 约 `127.4 MiB`，均低于 ceiling；`lohPeakBytes` 以测量点仍被引用的大对象负载与 GC LOH 存量的较大值记录，不以强制 GC 后的零值替代场景峰值。

## 当前已知风险

- Unique 已移除 CSV 每行 `CloneDuplicateValues` 全量复制，改为 committed state + pending journal；BenchmarkDotNet 与 Round 7 独立资源探针均覆盖 100K 行。
- 多规则、动态 Plan Build、租户缓存、JSON/XML v1/v2 parse 和 assembly scan/explicit registration 已有 ShortRun 结果；Round 7 已补齐受控进程 LOH/Peak Working Set ceiling 证据。

## 必须记录

- Wall Time、Operations/sec、Allocated Bytes。
- Gen0、Gen1、Gen2、LOH、Peak Working Set。
- 10K/100K 行、1/5 Unique 列、10K 多规则、100/500 动态列 Plan Build、100/1000 租户缓存、JSON/XML v1/v2、Assembly scan/explicit registration。

不得宣称 0GC，除非对应 Benchmark 明确证明热路径没有分配。

## Review Fix Round 8 资源证据

命令：
`dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build -- --resource-probe artifacts/mapval-v2/resource-probe-round8.jsonl`

本轮修复了 Round 7 资源探针的两个证据缺口：每个场景先实际构建 `tenantCount` 个不同租户的 Plan，再执行 `planBuildCount` 次额外构建；子进程结果以结构化 JSON 对象写入 JSONL，不再把场景结果存为转义字符串。LOH 记录拆分为 `lohSampledPeakBytes`（计划、Unique 和 payload 阶段采样的 `GenerationInfo[3].SizeBeforeBytes` 与 retained 值的最大值）和 `lohRetainedBytes`（强制 GC 后仍存活的 payload/LOH 存量）。

环境：Windows，X64，`.NET 8.0.27`，Release，进程路径为 `benchmarks/Bing.Offices.Benchmarks/bin/Release/net8.0/Bing.Offices.Benchmarks.exe`。ceiling 为 `LOH <= 512 MiB`、`Peak Working Set <= 1024 MiB`。

原始 artifact：`artifacts/mapval-v2/resource-probe-round8.jsonl`。独立解析结果：

- child 场景数：`16`；所有 `exitCode=0`、`status=passed`。
- `tenantCount=1000` 场景的每条记录均为 `tenantPlanCount=1000`。
- 最大 `lohSampledPeakBytes`：`85,801,616` bytes。
- 最大 `lohRetainedBytes`：`85,801,616` bytes。
- 最大 `peakWorkingSetBytes`：`133,238,784` bytes。
- 资源 ceiling 判定：全部通过。

JSONL header 说明采样方法，child 的 `result` 字段可直接作为 JSON 对象解析；不再依赖人工解码嵌套字符串。该证据仍是受控子进程中的阶段采样，不等同于操作系统级全程采样 profiler，但已覆盖计划声明的 tenant、PlanBuild、Unique 列/行组合和资源 ceiling。
