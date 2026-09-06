# Resource 报告

## 结论

状态：`PARTIAL / UNAPPROVED`。本轮真实运行 Excel 14 个独立进程场景、Mapping/Unique 16 个子进程场景和 4 档并发尾延迟；全部执行成功或按预期结构化拒绝。没有维护者批准的资源/延迟预算，未完成 1k/10k/100k Excel 与 Failure Workbook 双 DOM 全矩阵，因此不能宣称低 GC 或资源门禁通过。

## 环境与原始产物

- Windows 10.0.19045 win-x64；SDK 10.0.400；Runtime .NET 8.0.30；Workstation GC。
- `artifacts/resource/excel-resource-probe-rerun.jsonl`
- `artifacts/resource/mapping-resource-probe.jsonl`
- `artifacts/resource/tail-latency.jsonl`

## Excel 独立进程探针

`Import_ResourceProbe_ShouldRunInIndependentProcess` 为 14/14。正常 XLSX、250x4 DOM、shared strings、402 styles、6 pictures 和 XLS/OLE 均成功；`dom-limit` 在 250 source rows 中只导入 100 行并返回结构化 resource-limit。ZIP total、compression ratio、shared strings、styles、worksheet、XML depth 和 XML characters 七个场景均在 `Preflight` 拒绝，`importedRows=0`。

最大观测 PeakWorkingSet 为 55,476,224 bytes（250x4 DOM）；OLE 为 42,639,360 bytes；Preflight 拒绝约 30.5–32.5 MB。该值是独立短生命周期进程的 `PeakWorkingSet64`，不是任意输入硬上限。

## Mapping/Unique 与 LOH

16/16 子进程通过。矩阵覆盖 plan build 100/500、tenant 100/1000、unique columns 1/5、unique rows 10,000/100,000。最大观测：LOH sampled/retained 56,800,952 bytes；PeakWorkingSet 137,912,320 bytes；最长场景 502.3165 ms。探针 ceiling 分别是 512 MiB LOH 与 1 GiB working set，但属于测试内置保护值，不是获批准的产品预算。

## 尾延迟

每档 5,000 样本，workload 为冷 Mapping plan build：

| 并发 | P50 us | P95 us | P99 us | 吞吐 ops/s |
| ---: | ---: | ---: | ---: | ---: |
| 1 | 8,768 | 19,055 | 20,709 | 49,297.61 |
| 4 | 7,240 | 18,498 | 21,328 | 51,277.74 |
| 16 | 10,382 | 36,246 | 42,422 | 30,398.32 |
| 64 | 41,309 | 106,115 | 119,164 | 9,329.44 |

## 未完成与解除条件

- Failure Workbook `ErrorRowsOnly` 1k/10k/100k 已完成 MemoryDiagnoser，但 `AnnotatedOriginal` / `ErrorRowsOnly` 双 DOM 的独立 PeakWorkingSet、LOH 保留量和序列化输出峰值尚未采集。
- 1k/10k/100k Excel Import/Export 与 CSV 已完成 MemoryDiagnoser；图片/模板/样式/验证的完整资源矩阵和独立进程 PeakWorkingSet 尚未执行。
- 取消延迟没有独立分位数；当前只有确定性取消正确性测试。
- 没有预算批准人和批准阈值。

解除条件：维护者提供性能/资源预算与可比 baseline，在固定机器运行完整 Benchmark/Resource 矩阵并复核 Failure Workbook 双 DOM；若结果超预算，优化后用同一 Job 和输入重跑。
