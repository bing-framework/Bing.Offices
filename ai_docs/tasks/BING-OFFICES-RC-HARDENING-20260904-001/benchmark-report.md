# Benchmark Report

## 当前任务结果

- Benchmark project：`benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj`
- 运行：Release，BenchmarkDotNet `0.14.0`，ShortRun，`LaunchCount=1`、`WarmupCount=3`、`IterationCount=3`。
- 环境：Windows 10 `10.0.19045.6466`，Intel Core Ultra 7 270K Plus，24 logical/physical cores，.NET 8.0.30，Concurrent Workstation GC。
- 结果：MappingValidation `10/10`，退出码 `0`。
- Markdown：`BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.MappingValidationBenchmarks-report-github.md`
- CSV/HTML：同名 results 目录。
- 完整日志：`BenchmarkDotNet.Artifacts/Bing.Offices.Benchmarks.MappingValidationBenchmarks-20260904-165240.log`

## 代表性结果

| 场景 | Mean | Allocated |
| --- | ---: | ---: |
| `ParseJsonV2` | `3,985.2 ns` | `10.25 KB` |
| `ParseXmlV2` | `12,247.5 ns` | `57.18 KB` |
| `MultiRulePlanBuild` | `247,346.1 ns` | `284.72 KB` |
| `CacheKeyStringToUtf8` | `2,325.1 ns` | `5.89 KB` |
| `CacheKeyUtf8Bytes` | `2,098.0 ns` | `2.52 KB` |
| `ExplicitRegistration` | `226.3 ns` | `1.29 KB` |
| `AssemblyScanRegistration` | `1,541.2 ns` | `3.45 KB` |

`PeakWorkingSetBytes` 仍是独立 probe 相关场景，不据此宣称 BDN 能给出进程峰值硬上限。

## 解释边界

- 本次是当前任务有效的 MappingValidation ShortRun 证据，不是正式性能通过。
- 没有已批准的吞吐、分配、尾延迟或工作集 budget，因此状态为 `PARTIAL / UNAPPROVED`。
- 当前任务没有完成完整 Excel/CSV 1k/10k/100k、Failure Workbook 双 DOM、全 StreamPipeline 和 CSV writer A/B 矩阵。
- 历史 StreamPipeline 100k Import 曾出现 3 个异常；未在当前任务中形成可接受的完整解释，不能作为当前性能结论。
- 历史 Failure Workbook 100k 分配约 2.5 GB，仅作为风险背景，不写入当前测量结果。

## 结论

当前 benchmark 执行有效且 `10/10` 完成；发布性能门禁 `UNAPPROVED`，不能给 Go。
