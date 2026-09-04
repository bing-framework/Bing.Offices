# Benchmark 执行计划

本报告仅在真实运行后填写结果，不以 Benchmark 方法存在代替证据。

## 目标场景

- CSV 导入/导出：1k、10k、100k。
- 编译 Getter/Setter 与反射对照。
- Mapping Plan cache 命中、未命中、并发、长配置键。
- Validation range、命名规则 10/100/1000/10000。
- Failure Workbook：1%/10%/100% 错误、XLS/XLSX、图片有无。
- `ExportToBytes`、Stream、File 峰值与分配。
- 并发 1/4/16/64、取消延迟、ArrayPool（仅有证据时）。

## 必备记录

环境、commit/diff、runtime/TFM、JIT/GC、包版本、参数、Mean/Error/StdDev/P95/P99、Allocated、Gen0/1/2、LOH、PeakWorkingSet、吞吐、异常/离群/OOM、原始 artifact 路径。BDN 进程峰值不与独立 ResourceProbe 峰值混用。

## 当前状态

P5-01：`PARTIAL`。`StreamPipelineBenchmarks` 的 Import/Export/ExportDestinationCapacity × 1k/10k/100k 已以 ShortRun 执行并完成 `9/9`；Round 8 最终编译版本已执行 `TenantPlanCache` 的命中/淘汰场景和 compiled getter/setter 与反射对照场景。MappingValidation、FailureWorkbook、DynamicPlan、UniqueJournal 的既有结果也已读取并纳入 `benchmark-report.md`。Round 8 使 `TenantPlanCacheEviction` 通过 Provider SPI 显式使用实际容量，并在容量小于租户数时断言重建行为。CSV 完整统一批次、完整 failure workbook 格式/图片矩阵、ExportToBytes/Stream/File、取消延迟和正式预算仍未完成。

本轮没有根据短基准结果实施未经对照的优化，也没有把 BDN 托管分配当作独立进程工作集。100,000 行 Import 的 `3` 个异常计数和宽置信区间保留在报告中。
