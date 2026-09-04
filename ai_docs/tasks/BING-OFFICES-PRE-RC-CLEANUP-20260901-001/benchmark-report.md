# Benchmark 执行报告

## 结论

本报告记录当前工作树可复核的真实 BenchmarkDotNet 与独立基准结果，不将 ShortRun、宽置信区间或未批准预算解释为发布性能门禁通过。当前性能结论为 `UNAPPROVED`，发布判定仍为 No-Go。

## Round 8 基准修复

- `TenantPlanCacheBenchmarks.TenantPlanCacheEviction()` 现在通过 `ExcelMappingPlanFactoryProvider.CreateDefault(cacheCapacity: capacity)` 创建与日志一致的实际缓存容量。
- 基准在 `TenantCount > capacity` 时直接断言首个租户计划会被淘汰并重建；net8 单元回归覆盖显式容量为 `2` 的 hit/miss/eviction 语义。
- Round 8 未将历史 TenantPlanCache artifact 重写为新测量；在重新运行该 Benchmark 前，历史结果只作为旧证据，不作为本轮容量修复后的数值结论。

## Round 8 最终编译版本补充基准

- 构建：`dotnet build .\Bing.Offices.sln -c Release --no-restore`，退出码 `0`，`0 error / 27 warning`。
- TenantPlanCache 命令：

  ```powershell
  $env:NUGET_PACKAGES = $null
  dotnet run --project .\benchmarks\Bing.Offices.Benchmarks\Bing.Offices.Benchmarks.csproj -c Release --no-build -- -j short -m 3 -f "*TenantPlanCacheBenchmarks*"
  ```
- TenantPlanCache 退出码：`0`；最终日志：`BenchmarkDotNet.Artifacts/Bing.Offices.Benchmarks.TenantPlanCacheBenchmarks-20260902-214530.log`。
- 最终日志重复输出均满足：`TenantCount=100` 为 `capacity=50`、`observed=True`；`TenantCount=1000` 为 `capacity=256`、`observed=True`。基准实际通过 `ExcelMappingPlanFactoryProvider.CreateDefault(cacheCapacity: capacity)` 创建 Factory，并对 `TenantCount > capacity` 断言首项计划被淘汰后重建。
- 最终结果 Markdown：`BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.TenantPlanCacheBenchmarks-report-github.md`。本批次数值为：

  | Method | TenantCount | Mean | Error | StdDev | Gen0 | Gen1 | Allocated |
  | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
  | TenantPlanCache | 100 | 350.9 μs | 27.54 μs | 1.51 μs | 27.3438 | 1.9531 | 505.77 KB |
  | TenantPlanCacheEviction | 100 | 423.9 μs | 78.01 μs | 4.28 μs | 27.3438 | 1.9531 | 509.41 KB |
  | TenantPlanCache | 1000 | 13,612.2 μs | 16,264.82 μs | 891.53 μs | 265.6250 | 93.7500 | 4,916.35 KB |
  | TenantPlanCacheEviction | 1000 | 13,659.4 μs | 15,648.38 μs | 857.74 μs | 265.6250 | 93.7500 | 4,924.58 KB |

- PropertyAccessor 命令：

  ```powershell
  $env:NUGET_PACKAGES = $null
  dotnet run --project .\benchmarks\Bing.Offices.Benchmarks\Bing.Offices.Benchmarks.csproj -c Release --no-build -- -j short -m 3 -f "*PropertyAccessorBenchmarks*"
  ```
- PropertyAccessor 退出码：`0`；日志：`BenchmarkDotNet.Artifacts/Bing.Offices.Benchmarks.PropertyAccessorBenchmarks-20260902-214303.log`；结果：`BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.PropertyAccessorBenchmarks-report-github.md`。
- 结果：compiled getter `0.9261 ns` / `0 B`，reflection getter `3.1814 ns` / `0 B`，compiled setter `1.3012 ns` / `0 B`，reflection setter `5.2867 ns` / `0 B`。该场景仅证明当前访问器实现的相对热路径样本，不构成完整导入/导出性能预算或发布门禁。

## 当前 StreamPipeline BenchmarkDotNet 运行

- 命令：

  ```powershell
  $env:NUGET_PACKAGES = $null
  dotnet run --project .\benchmarks\Bing.Offices.Benchmarks\Bing.Offices.Benchmarks.csproj -c Release --no-build -- -j short -m 3 -f "*StreamPipelineBenchmarks*"
  ```
- 退出码：`0`
- 场景数：`9/9` 完成。
- 运行时：`.NET 8.0.27 (8.0.2726.22922)`。
- OS：Windows 10 `10.0.19045.6466` / 22H2。
- CPU：Intel Core Ultra 7 270K Plus，24 logical / 24 physical cores，X64，AVX2。
- SDK：`.NET SDK 10.0.300`。
- GC：Concurrent Workstation。
- BenchmarkDotNet：`0.14.0`。
- Job：`ShortRun`，`LaunchCount=1`，`WarmupCount=3`，`IterationCount=3`。
- 运行日志：`BenchmarkDotNet.Artifacts/Bing.Offices.Benchmarks.StreamPipelineBenchmarks-20260901-214606.log`。
- 原始结果：
  - [CSV 报告](../../BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.StreamPipelineBenchmarks-report.csv)
  - [Markdown 报告](../../BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.StreamPipelineBenchmarks-report-github.md)
  - [HTML 报告](../../BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.StreamPipelineBenchmarks-report.html)

第一次运行在 BenchmarkDotNet 自动生成项目阶段失败，原因是环境变量 `$env:NUGET_PACKAGES` 指向了隔离 package consumer 的不完整缓存；未产生可用测量。清除该环境变量后按上方命令重跑成功。该环境修复没有修改测试断言或基准逻辑。

## StreamPipeline 结果

| Method | RowCount | Mean | Error | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Import | 1,000 | 29.91 ms | 12.25 ms | 0.671 ms | 1,100.0000 | 500.0000 | 100.0000 | 17.14 MB |
| Export | 1,000 | 16.86 ms | 67.31 ms | 3.690 ms | 625.0000 | 406.2500 | 218.7500 | 8.39 MB |
| ExportDestinationCapacity | 1,000 | 15.51 ms | 87.48 ms | 4.795 ms | 625.0000 | 406.2500 | 218.7500 | 8.39 MB |
| Import | 10,000 | 292.31 ms | 693.48 ms | 38.012 ms | 11,000.0000 | 7,000.0000 | 2,000.0000 | 162.85 MB |
| Export | 10,000 | 120.77 ms | 42.18 ms | 2.312 ms | 4,666.6667 | 3,333.3333 | 1,333.3333 | 71.25 MB |
| ExportDestinationCapacity | 10,000 | 129.56 ms | 362.93 ms | 19.894 ms | 4,666.6667 | 3,333.3333 | 1,333.3333 | 71.25 MB |
| Import | 100,000 | 2,492.64 ms | 8,943.24 ms | 490.209 ms | 93,000.0000 | 48,000.0000 | 2,000.0000 | 1,623.15 MB |
| Export | 100,000 | 1,396.48 ms | 366.71 ms | 20.100 ms | 38,000.0000 | 24,000.0000 | 4,000.0000 | 734.86 MB |
| ExportDestinationCapacity | 100,000 | 1,392.61 ms | 1,093.66 ms | 59.947 ms | 38,000.0000 | 24,000.0000 | 4,000.0000 | 734.86 MB |

100,000 行 Import 的基准进程记录了 `3` 个异常，但进程退出码为 `0`；该异常计数必须作为结果风险保留，不能从报告中删除或按成功吞并。

BDN 结果只提供样本均值、误差、标准差和托管分配，不提供本轮的 P95/P99、LOH 或独立进程工作集，因此不能替代 [resource-report.md](resource-report.md) 中的 ResourceProbe/尾延迟证据。

## 已存在的补充标准基准产物

以下结果文件已存在于当前工作树，作为补充取证；它们不是本次 9 场景 StreamPipeline ShortRun 的替代品，且在没有统一重新运行时间、提交身份和批准预算前不作为发布性能通过依据：

| 基准组 | 代表结果 |
| --- | --- |
| MappingValidation | ParseJson/Xml、v1/v2、计划构建、缓存键和注册分支；见 `BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.MappingValidationBenchmarks-report-github.md` |
| FailureWorkbook | 1,000 / 10,000 / 100,000 failure rows；100,000 行约 `4,712.89 ms`、`2,554.11 MB`；失败输出限制现称 `MaxSerializedBytes`；见 `BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.FailureWorkbookBenchmarks-report-github.md` |
| DynamicPlan | 冷构建、缓存命中和缓存未命中；见 `BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.DynamicPlanBenchmarks-report-github.md` |
| UniqueJournal | 1/5 列与 10,000/100,000 行；见 `BenchmarkDotNet.Artifacts/results/Bing.Offices.Benchmarks.UniqueJournalBenchmarks-report-github.md` |
| TenantPlanCache | Round 8 最终运行日志 `BenchmarkDotNet.Artifacts/Bing.Offices.Benchmarks.TenantPlanCacheBenchmarks-20260902-214530.log` 已出现 `capacity=50`（100 tenant）和 `capacity=256`（1,000 tenant），且每次 `observed=True`；最终结果 Markdown 已重新导出并记录本批次数值。 |

## 尚未完成或不能判定的性能项

- CSV 1k/10k/100k 的本轮统一 BDN 结果未形成独立、同提交的完整矩阵。
- 验证规则 10/100/1,000/10,000、失败工作簿 XLS/XLSX/图片分支、ExportToBytes/Stream/File 和取消延迟尚未形成同一运行批次的完整报告；compiled getter/setter 与反射对照已形成 Round 8 独立 ShortRun，但不等同于完整矩阵。
- 并发 1/4/16/64 的尾延迟已测量，但预算未批准，不能判定门禁；详见 `resource-report.md`。
- 未进行基于结果的优化；没有把 Span、ArrayPool、缓存或其它技术选择表述为零分配或性能通过。

## 性能判定

状态：`PARTIAL / UNAPPROVED`。

解除条件：统一提交身份与环境下补齐计划要求的代表场景，维护者批准明确的时间、分配、LOH、工作集和尾延迟预算，复核 100,000 行 Import 的异常计数，并完成独立 Review。