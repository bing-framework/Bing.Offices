# Benchmark Report

- Task ID：`BING-OFFICES-POST-MINIEXCEL-HARDENING-20260916-001`
- 配置：`.NET 8.0.30`，Windows x64 本机。

## 入口和目录

BenchmarkDotNet 使用 `ManualConfig.WithArtifactsPath`，正式目录为 `artifacts/benchmarks/`。基准列表命令通过，并包含 `MiniExcelRealIoBenchmarks.ExportImportSync` 和 `ExportImportAsync`；列表证据为 `artifacts/benchmarks/list-flat.txt`。

## 100K controlled probe

证据：`artifacts/benchmarks/miniexcel-probe-100k.json`，schema 2。

| 路径 | elapsed | allocated | output | rows/sec | Gen0/1/2 | working set / peak |
| --- | ---: | ---: | ---: | ---: | --- | ---: |
| Sync | 1972.34 ms | 1,276,698,832 B | 2,078,325 B | 50,701 | 102/13/3 | 111,501,312 / 111,611,904 B |
| Async | 1258.70 ms | 1,275,428,960 B | 2,078,320 B | 79,447 | 100/11/1 | 133,664,768 / 133,664,768 B |

与本任务较早 probe 记录相比，当前优化后的 allocation 和 output 已重新采集；elapsed、GC 和 working set 受机器状态影响，不能当作跨运行稳定基线。优化点是缓存固定列反射绑定、动态 key 集合和增加可审计指标，没有引入未经证明的池化或并发。

## Relation complexity

MiniExcel relation binder 保持按 parent 数组逐 child `FirstOrDefault` 的最坏 `O(P×C)`，本任务没有在缺少稳定 benchmark 证据时改写 comparer/重复键语义。

## 未验证

500K、1M、生产 2C/4GiB 机器和外部 CI 均为 `NOT_VERIFIED`。本地 100K 结果不外推到这些容量或环境。

## NPOI/MiniExcel 同 workload 对照

新增 `ProviderComparisonBenchmarks`（BenchmarkDotNet）和受控 probe，使用同一批 100K 行、同一 Workbook request，分别执行 Export+Import roundtrip 的 Sync/Async 路径。原始样本位于 `artifacts/review-fix/provider-comparison-100k-final.json`，包含两 Provider、两模式各 3 次重复；`phase=after` 表示当前工作树测量。

| Provider / 模式 | 中位 elapsed | 中位 allocated | 中位 output | 中位 rows/sec |
| --- | ---: | ---: | ---: | ---: |
| NPOI Sync | 4991.10 ms | 2,164,451,024 B | 2,186,768 B | 20,036 |
| NPOI Async | 5176.89 ms | 2,164,807,400 B | 2,186,768 B | 19,317 |
| MiniExcel Sync | 1358.25 ms | 1,278,084,024 B | 2,099,675 B | 73,624 |
| MiniExcel Async | 1796.84 ms | 1,278,210,104 B | 2,099,669 B | 55,653 |

这组数据只证明当前环境下的同 workload 对照，不构成优化前后收益结论：本任务没有保留可重放的修改前二进制/提交基线，因此 `before` 数据为 `NOT_VERIFIED`，不得由当前 `after` 样本推导收益。需要发布级性能结论时，应在固定环境从前一 revision 运行相同 probe，并与本 artifact 逐字段比较。
