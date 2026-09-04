# 资源与尾延迟报告

## 结论

本报告分开记录两套不同的独立子进程 workload：真实 XLS/XLSX 导入的 Excel ResourceProbe，以及 mapping-plan/unique-tracker 资源探针。二者的输入、输出字段和验收目标不同，不交叉推导结论。尾延迟结果另行记录，`budgetStatus=UNAPPROVED`，不能作为发布资源或性能门禁通过。

## Excel ResourceProbe

### 运行证据

- 原始 JSONL：[excel-resource-probe-rerun.jsonl](artifacts/excel-resource-probe-rerun.jsonl)
- 启动测试：`Bing.Offices.Tests.ExcelP0RegressionTest.Import_ResourceProbe_ShouldRunInIndependentProcess`
- 子进程程序：`tests/Bing.Offices.ResourceProbe/Program.cs`，每个模式均由独立 `dotnet Bing.Offices.ResourceProbe.dll` 进程执行
- 模式：`zip`、`dom`、`dom-limit`、`shared-strings`、`styles`、`drawings`、`ole`
- Artifact 记录：`scenario`、`inputSha256`、`inputBytes`、`mode`、`status`、`exitCode`、`sheets`、`rows`、`columns`、`cells`、`importedRows`、`sharedStrings`、`styles`、`pictures`、`elapsedMs`、`peakWorkingSet`
- 记录数：`7`；全部 `exitCode=0`；最大 `peakWorkingSet=53,276,672` bytes；最大 `rows=250`

### 七模式结果

| Scenario | Input bytes | Rows | Imported rows | Other metrics | Status | Exit code | Peak working set |
| --- | ---: | ---: | ---: | --- | --- | ---: | ---: |
| `zip` | 4175 | 4 | 3 | sheets=1, columns=2, cells=8, sharedStrings=8, styles=2, pictures=0 | `success` | 0 | 50,106,368 |
| `dom` | 11339 | 250 | 249 | sheets=1, columns=4, cells=1000, sharedStrings=1000, styles=2, pictures=0 | `success` | 0 | 53,276,672 |
| `dom-limit` | 11339 | 250 | 100 | sheets=1, columns=4, cells=1000, sharedStrings=1000, styles=2, pictures=0 | `resource-limit` | 0 | 52,953,088 |
| `shared-strings` | 6540 | 20 | 19 | sheets=1, columns=20, cells=400, sharedStrings=400, styles=2, pictures=0 | `success` | 0 | 51,163,136 |
| `styles` | 7782 | 20 | 19 | sheets=1, columns=20, cells=400, sharedStrings=400, styles=402, pictures=0 | `success` | 0 | 51,847,168 |
| `drawings` | 6872 | 10 | 9 | sheets=1, columns=4, cells=40, sharedStrings=40, styles=2, pictures=6 | `success` | 0 | 50,933,760 |
| `ole` | 4096 | -1 | 3 | XLS/OLE preflight metrics unavailable; sheets/rows/columns/cells/sharedStrings/styles/pictures=-1 | `success` | 0 | 40,583,168 |

`dom-limit` 的 `rows=250` 是输入 Workbook 的预解析行数，`importedRows=100` 是资源限制后的实际导入行数；该区别用于证明限制生效，不是计数冲突。

## Mapping/Unique ResourceProbe

- 原始 JSONL：[resource-probe-rerun.jsonl](artifacts/resource-probe-rerun.jsonl)
- 运行时间：`2026-09-01T13:30:35.8985293Z` 起。
- 探针进程：`benchmarks/Bing.Offices.Benchmarks/bin/Release/net8.0/Bing.Offices.Benchmarks.exe`。
- Runtime：`.NET 8.0.27`。
- LOH ceiling：`536,870,912` bytes（512 MiB）。
- Peak working set ceiling：`1,073,741,824` bytes（1 GiB）。
- 子进程结果：`16/16` 的 `exitCode=0`，场景 `status=passed`；这是 mapping-plan/unique-tracker workload，不是 Excel ResourceProbe。
- 场景维度：`planBuildCount`、`tenantCount`、`uniqueColumnCount`、`uniqueRowCount`；覆盖 10,000 与 100,000 行、1 与 5 个唯一列、100 与 1,000 tenant、100 与 500 plan builds。

## 观测上限与峰值

| 指标 | 观测最大值 | 对应含义 |
| --- | ---: | --- |
| `lohSampledPeakBytes` | 57,147,216 bytes（约 54.48 MiB） | 每个 workload phase 后采样到的 Generation 3 `SizeBeforeBytes` 最大值 |
| `lohRetainedBytes` | 57,147,216 bytes（约 54.48 MiB） | 强制 GC 后仍保留的 live payload 最大值 |
| `peakWorkingSetBytes` | 141,549,568 bytes（约 135.00 MiB） | 子进程工作集最大值 |
| LOH ceiling | 536,870,912 bytes | 探针配置阈值，不是产品完整内存上限 |
| Working set ceiling | 1,073,741,824 bytes | 探针配置阈值，不是产品完整内存上限 |

所有样本低于探针配置阈值，但该结论仅适用于当前场景、当前 runtime、当前输入生成方式和当前 Windows 进程测量；不代表 NPOI DOM 对任意压缩包、工作簿、实体集合或失败工作簿都有硬内存保证。

## 探针污染控制

Excel ResourceProbe 的测量路径由独立 child scenario 承担，不再在导入前由父进程执行 `WorkbookFactory.Create()`。因此父进程没有预先构造与被测导入相同的 Workbook DOM，峰值来源可以归因于 Excel child workload 生命周期。Mapping/Unique 探针同样以独立 benchmark child scenario 测量，但其峰值只适用于自身的 plan/tracker 场景。两套证据都不能将 NPOI DOM 的解压、DOM、业务实体和失败输出峰值合并为一个由 `MaxInputBytes` 保证的上限。

## 资源边界

- `MaxInputBytes` 及相关输入限制不能单独控制 ZIP 解压内容、NPOI DOM、业务实体、图片对象和 Failure Workbook 的峰值。
- Excel ResourceProbe 证明七个指定 XLS/XLSX 样本可由独立进程执行并记录输入维度、导入结果和工作集；Mapping/Unique ResourceProbe 证明其指定 plan/tracker 场景的工作集与 LOH 行为。二者都不是任意 XLS/XLSX 输入的安全证明。
- Failure Workbook 可能同时保留原始 Workbook、错误行数据和输出 Workbook，存在双 DOM/多对象图风险；其结果需要独立 failure-workbook 矩阵，不能用本报告的 mapping 场景替代。
- 本报告未把 BDN 进程的托管分配与独立 Probe 工作集混合计算。

## 未覆盖边界

- 当前 Excel artifact 不覆盖 Failure Workbook 双 DOM、压缩比放大、取消延迟或任意未生成的恶意/超大输入。
- XLS/OLE 样本的 ZIP 元数据字段不可用，报告保留 `-1`，不把不可测字段写成零或成功预检。
- 因此不能把 `MaxInputBytes`、本次最大工作集或任一 ceiling 写成任意输入的硬内存上限；Failure Workbook 双 DOM 仍需独立矩阵，当前按发布阻断保留。

## 尾延迟运行证据

- 原始 JSONL：[tail-latency-rerun.jsonl](artifacts/tail-latency-rerun.jsonl)
- Runtime：`.NET 8.0.27`。
- OS：Windows 10 `10.0.19045`。
- Processor count：`24`。
- Workload：`cold-plan-build`。
- Warmup：`64` operations；每个并发度 `256` operations × `5` repetitions；总样本每个场景 `1,280`。
- 延迟定义：从队列提交时间戳到 mapping plan 完成的端到端样本。
- `budgetStatus`：所有记录均为 `UNAPPROVED`。

| Concurrency | P50 | P95 | P99 | Throughput |
| ---: | ---: | ---: | ---: | ---: |
| 1 | 1,861 μs | 3,621 μs | 4,257 μs | 67,611.82 op/s |
| 4 | 667 μs | 2,961 μs | 3,230 μs | 145,807.47 op/s |
| 16 | 699 μs | 2,768 μs | 3,268 μs | 133,587.99 op/s |
| 64 | 1,711 μs | 6,169 μs | 6,902 μs | 52,305.10 op/s |

尾延迟的重复样本存在明显波动，例如并发 64 的单次重复 P99 达到 `7,116 μs`，并发 16 的单次重复 P99 达到 `3,352 μs`。在没有批准的预算、稳定性目标和同提交对照基线之前，不作通过或回归结论。

## 状态与解除条件

状态：`VERIFIED`（探针执行与数据完整性）+ `UNAPPROVED`（资源/尾延迟门禁）。

解除条件：补齐真实 XLS/XLSX、Failure Workbook、输入压缩包/解压、取消和并发场景的独立子进程证据；明确并批准预算；完成双 DOM 风险审计、文档边界同步和独立 Review。