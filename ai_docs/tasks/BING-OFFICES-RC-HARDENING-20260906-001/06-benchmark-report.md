# Benchmark 与资源报告

## 已执行

- Benchmark Release 构建通过，当前 solution build 为 `0 warning / 0 error`。
- 资源探针：`artifacts/resource-scenario-before.jsonl` 与 `artifacts/resource-scenario-candidate.jsonl`，10 plan builds / 10 tenants / 100 unique rows，两侧 status 均为 `passed`，LOH sampled peak 均为 `0` bytes。
- 取消尾延迟探针：`artifacts/tail-latency.jsonl`，1/4/16/64 并发、10 操作、P50/P95/P99 原始样本，状态 `measured`。

## Before/Candidate 对比

同一参数 `planBuildCount=10, tenantCount=10, uniqueColumnCount=1, uniqueRowCount=100`，分别使用基线 commit 临时副本和候选 Release benchmark：

| 指标 | Before | Candidate | 变化 |
| --- | ---: | ---: | ---: |
| PeakWorkingSet | 36,278,272 | 36,294,656 | +16,384（约 +0.05%） |
| LOH sampled peak | 0 | 0 | 0 |
| Elapsed ms | 109.1121 | 108.2381 | -0.80% |
| Status | passed | passed | 预算仍未批准 |

原始文件：`artifacts/resource-scenario-before.jsonl`、`artifacts/resource-scenario-candidate.jsonl`。

CSV 1M 命令：`dotnet restore benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj --configfile <task>/artifacts/benchmark/NuGet.offline.config --ignore-failed-sources`；`dotnet build ... -c Release --no-restore -t:Rebuild`；`dotnet run --project benchmarks/Bing.Offices.Benchmarks/Bing.Offices.Benchmarks.csproj -c Release --no-build --no-restore -- --filter '*CsvPipelineBenchmarks*' --job short --memory --inProcess --artifacts <task>/artifacts/benchmark/csv-pipeline-1m-inprocess-final --exporters json markdown csv`。

## 尚未满足

- 未采集同机同 Job 的 before/candidate 完整 1k/10k/100k/1M 矩阵。
- 未完成图片/模板/图表/Failure Workbook 双 DOM 峰值矩阵和跨平台矩阵。
- 资源预算和性能回归阈值仍是 `UNAPPROVED`，因此不能判定 PASS。
- 未声称 Excel 真正流式、低 GC 或零分配。

## 2026-09-07 候选容量矩阵

使用当前 Release 代码、BenchmarkDotNet `ShortRun + MemoryDiagnoser + InProcessEmitToolchain` 完成候选侧矩阵。由于隔离 BDN toolchain 仍受 NuGet TLS 阻断，这些结果不与隔离进程或历史结果计算回归比率。

| 场景 | 参数 | Mean | Allocated | 原始 artifact |
| --- | ---: | ---: | ---: | --- |
| Excel Import | 1k / 10k / 100k | 30.139 ms / 211.231 ms / 2.093 s | 17.86 MB / 168.75 MB / 1,684.76 MB | `artifacts/benchmark/current-stream-pipeline-inprocess` |
| Excel Export | 1k / 10k / 100k | 10.988 ms / 123.226 ms / 1.371 s | 8.37 MB / 71.13 MB / 734.53 MB | `artifacts/benchmark/current-stream-pipeline-inprocess` |
| Excel ExportDestinationCapacity | 1k / 10k / 100k | 9.427 ms / 116.829 ms / 1.340 s | 8.37 MB / 71.13 MB / 734.53 MB | `artifacts/benchmark/current-stream-pipeline-inprocess` |
| Failure Workbook ErrorRowsOnly | 1k / 10k / 100k | 68.910 ms / 367.633 ms / 4.310 s | 29.02 MB / 263.89 MB / 2,606.08 MB | `artifacts/benchmark/current-failure-workbook-inprocess` |
| HeaderStyle | 1k / 10k / 100k | 53.390 ms / 507.610 ms / 5.190 s | 29.11 MB / 281.90 MB / 2,777.76 MB | `artifacts/benchmark/current-header-validation-inprocess` |
| ValidationRange | 1k / 10k / 100k | 4.160 ms / 61.490 ms / 662.847 ms | 5.47 MB / 49.18 MB / 538.74 MB | `artifacts/benchmark/current-header-validation-inprocess` |

上述候选矩阵与 CSV 1M 结果合计覆盖当前计划要求的主要行数规模；仍不能证明 before/candidate 回归阈值、图片/模板/图表峰值或业务预算。

## CSV 1M 续跑

为覆盖计划要求的 CSV `1_000_000` 行场景，新增 `[Params(1000, 10000, 100000, 1_000_000)]` 并运行 8 个 Import/Export ShortRun 场景。由于 BenchmarkDotNet 隔离执行项目受 NuGet TLS 阻断，本次使用已构建程序集的 `InProcessEmitToolchain`；不与隔离进程结果混合做回归比率。

| Method | Rows | Mean | StdDev | Gen0 | Gen1 | Gen2 | Allocated |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Import | 1,000,000 | 945.480 ms | 4.871 ms | 110,000 | 29,000 | 1,000 | 1,984.1 MB |
| Export | 1,000,000 | 1.216 s | 3.445 ms | 798,000 | 1,000 | - | 14,455.7 MB |

1M 原始 artifact：`artifacts/benchmark/csv-pipeline-1m-inprocess-final/`；当前候选 1k/10k/100k 表和 Failure/Style/Validation 原始结果分别保留在 `artifacts/benchmark/current-*`。closure 目录中的历史结果仅作历史参考；所有预算依旧 `UNAPPROVED`。
