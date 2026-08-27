<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-RELEASE-HARDENING-20260825-01
AI_REVIEWED_AT: 2026-08-26T21:57:07.6015029+08:00

# Review Fix Round 10 独立复审报告

## 1. 验收摘要

本次复审以当前 `plan.md`、Round 10 `execution.md`、实际源码、Git Diff、BenchmarkDotNet 原始产物、资源探针、互操作产物和独立测试为证据，优先验证上一轮 `FIX-004`、`FIX-006`、`FIX-007`。

最终结论：**NEEDS_FIX**，继续保持 **NO-RELEASE**。

- `FIX-001` 至 `FIX-003`：`RESOLVED`，既有 P0 正确性合同未发现回归。
- `FIX-004`：`PARTIAL/BLOCKED`。GenerationId、manifest SHA 和四个 source SHA 仍一致，但 Excel COM 四次 `Workbooks.Open` 均失败，没有 roundtrip、截图、roundtrip SHA 或 NPOI 二次结构 PASS；LibreOffice/WPS 仍不可用。
- `FIX-005`：`RESOLVED`。Regex 容量、FIFO、timeout 和并发测试继续通过。
- `FIX-006`：`PARTIAL`。XLSX/XLS HeaderAttribute 双格式测试、BenchmarkDotNet 文件和真实 workload 资源探针已补充；但多个 benchmark 的参数标签没有进入实际 workload，style benchmark 没有使用 HeaderAttribute，validation benchmark 没有覆盖 ValidationRangeIndex 规模，资源探针的 retained 指标没有保持 UniqueTracker workload 存活。
- `FIX-007`：`RESOLVED`。迁移示例已改为根 builder 调用 `Build()`，`nuget-migration.md` 已纳入 fence 编译执行测试；未批准 breaking change 保持 DEFERRED，符合计划边界。

## 2. Review 边界与 Git 分析

- 当前已跟踪修改集中在发布加固涉及的生产代码、测试、文档、benchmark 和任务状态；任务目录及互操作/benchmark 证据仍包含未跟踪内容。
- Round 10 新增行为主要位于 `StreamPipelineBenchmarks`、`MappingValidationBenchmarks`、`Program.ResourceProbe`、HeaderAttribute 双格式测试和迁移文档 fence 测试。
- 未发现新增 public API、程序集依赖反转、NPOI 类型泄漏或第二套生产实现。
- `git diff --check` 通过，无空白错误；仅有 CRLF/LF 转换提示。
- Reviewer 未修改业务代码、测试代码、`plan.md` 或 `execution.md`；仅更新本文件。

## 3. 上一轮 FIX 复审矩阵

| FIX | Round 10 声明 | 复审状态 | 结论 |
| --- | --- | --- | --- |
| FIX-004 | PARTIAL/BLOCKED | PARTIAL/BLOCKED | 代际和 source SHA 可靠，真实办公客户端 roundtrip 仍不存在。 |
| FIX-005 | RESOLVED | RESOLVED | 专项与全量测试继续通过。 |
| FIX-006 | COMPLETED（本轮可执行范围） | PARTIAL | 双格式测试和产物已补，但部分 benchmark workload/参数及 retained 证据不成立。 |
| FIX-007 | COMPLETED/DEFERRED | RESOLVED | 文档缺陷和 fence 覆盖已修复；未批准 breaking change 合理延期。 |

## 4. 主要发现

### HIGH-001：真实办公客户端 roundtrip 门禁仍未完成

- `product-input-manifest.json`、`interop-dotnet-com.json` 和 `npoi-roundtrip-parse.json` 的 `GenerationId` 均为 `3eb6f7c1ae9f4afb8e245d7518b66b6f`。
- manifest SHA 与 COM `SourceManifestSha256` 一致；四个 source SHA 与 manifest 全部一致。
- Excel `16.0` Build `17932` 对四个当前 source 均记录独立尝试，但四条结果都是 `BLOCKED`。
- NPOI verifier 完整列出四个缺失 roundtrip 文件；没有截图、roundtrip SHA 或结构 PASS。
- LibreOffice/WPS 当前仍无可运行客户端证据。

因此 `RH-602` 继续是 Release blocker。当前证据证明验证过程没有混用旧代际，但不能证明产品文件经任一目标办公客户端保存重开后兼容。

### MEDIUM-001：性能产物存在，但 workload 与输入标签不一致

- `StreamPipelineBenchmarks` 对所有方法统一应用 `RowCount=1000/10000/100000`。
- `FailureWorkbookExport` 的 `_failureBytes` 始终只有表头和一个失败数据行，未使用 `RowCount`。原始报告三组 allocated 分别约 `504.49/505.02/503.77 KB`，说明三种参数标签对应相同规模 workload。
- `ValidationRange` 的 `_validationBytes` 始终只有一个数据行和一个单单元格 validation，未使用 `RowCount`，也没有直接构造大范围、重叠或多规则 `ValidationRangeIndex` workload。三组 allocated 约 `514.99/515.85/514.67 KB`。
- `HeaderStyle` 使用 `_rows.Take(Math.Min(RowCount, 1000))`，所以 `RowCount=10000/100000` 实际仍只处理 1000 行；原始报告对应 allocated 约 `8.98/8.72 MB`，不代表 10K/100K style workload。
- `HeaderStyle` 的模型 `BenchmarkRow` 没有 `HeaderAttribute`，该 benchmark 只测请求级 `HeaderStyle`，不能作为 `NpoiStyleCache.ApplyHeaderAttribute` 的性能证据。
- `MappingValidationBenchmarks` 的类级四组 Params 被应用到所有方法，导致 `RegexCacheHit` 生成 16 组与 Regex workload 无关的参数组合。结果可证明缓存命中的纳秒级绝对成本，但参数表会误导输入规模解释。

这些产物包含 Mean、Error、StdDev、GC、Allocated、CPU、SDK 和 runtime，文件格式本身合格；问题在于 workload 定义无法支撑 execution.md 声明的 failure/style/validation 大规模分组结论。

### MEDIUM-002：资源探针的 retained 指标未覆盖完整生产 workload

- 人工 `new byte[90 * 1024]` payload 已移除，16 个子进程均使用 mapping-plan 和 UniqueTracker 生产类型，改进成立。
- `RunScenario` 在强制 GC 前保持 `plans` 存活，并在末尾调用 `GC.KeepAlive(plans)`。
- UniqueTracker 及其 `values` 字典没有对应 `GC.KeepAlive`，最后一次使用发生在强制 GC 之前；JIT 可在 GC 点把它们视为不可达。
- 产物中所有 `lohRetainedBytes` 均为 `0`，因此该字段不能证明 UniqueTracker workload 的 retained capacity；它最多证明强制 GC 后当时仍被视为存活的对象未保留 LOH。
- Stream export 使用 fresh destination，解决了复用 `MemoryStream.SetLength(0)` 隐藏容量的问题，但当前 benchmark/资源产物仍没有记录 destination `Capacity` 或 stream retained-capacity 指标。

`artifacts/resource-round10.json` 可作为 mapping/unique 峰值工作集和 sampled LOH 的绝对证据，但不能按当前实现宣称完整 workload 的 retained capacity 已验证。

### LOW-001：execution.md 历史机器元数据仍重复

文件顶部是 Round 10 的严格终态三行，但正文中仍保留 Round 6 的第二组 `AI_EXECUTION_STATUS/AI_TASK_ID/AI_EXECUTION_FINISHED_AT`，Round 7 也只有汇总。该问题不改变代码结论，但继续降低历史自动解析可靠性；本轮不新增修复任务，避免把历史报告整理扩大为业务整改。

## 5. 已解决项

### FIX-005 RESOLVED

- Regex cache 专项与 net6/net8 全量测试继续通过。
- 容量 256、FIFO、重复命中、timeout 和并发上界合同未发现回归。

### FIX-007 RESOLVED

- `nuget-migration.md` 使用根 `mappingBuilder` 调用 `Build()`，不再把 `Build()` 链到列 builder。
- `DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually` 已枚举 `nuget-migration.md`，当前共提取并执行 10 个 C# fence。
- 独立专项验证中迁移 fence/当前 major consumer 为 2/2 PASS，Docs 全量为 9/9 PASS。
- breaking table 仍标记“待批准”，未删除或重命名 public API，符合 RH-201/RH-202 的批准前边界。

## 6. 结构化修复任务

### FIX-004

- 严重程度：HIGH
- 处理要求：MUST_FIX
- 当前状态：OPEN
- 复审状态：PARTIAL/BLOCKED
- 对应计划项：RH-602/RH-603
- 涉及文件/符号：`artifacts/interop-round2/Program.cs`、`artifacts/interop-round5/excel`、`execution.md`
- 问题：当前 Excel COM 验证对四个 source 均无法调用 `Workbooks.Open`，真实办公客户端保存重开矩阵仍为空。
- 证据：COM 四条结果均为 BLOCKED；四个 roundtrip、截图和 roundtrip SHA 不存在；NPOI verifier 列出四个缺失文件；LibreOffice/WPS 不可用。
- 影响：不能证明 XLSX/XLS 在 Excel、LibreOffice、WPS 保存重开后保留样式/ARGB、批注、合并、公式、验证、图片和失败工作簿结构。
- 修复目标：在可工作的真实客户端宿主中完成同一代际、哈希绑定的互操作矩阵；不可用客户端保留明确阻塞证据。
- 明确修复要求：在能够正常调用 Excel `Workbooks.Open` 的交互会话、位数和 Office 注册环境中，对同一或全新单一 `GenerationId` 的四个 source 完成打开、保存、重开、截图、roundtrip SHA 和 NPOI verifier 结构 PASS。LibreOffice/WPS 应记录版本并执行对应 XLSX/XLS 矩阵，或继续提供明确环境阻塞证据。不得复用旧代际 PASS 产物。
- 修复后的验证方式：Reviewer 核对 GenerationId、manifest/source/roundtrip SHA、客户端版本、时间顺序、四个客户端结果、截图、NPOI 结构断言及 source 前后 SHA 不变。

### FIX-006

- 严重程度：MEDIUM
- 处理要求：SHOULD_FIX
- 当前状态：OPEN
- 复审状态：PARTIAL
- 对应计划项：RH-402/RH-403/RH-404
- 涉及文件/符号：`StreamPipelineBenchmarks.Setup`、`FailureWorkbookExport`、`HeaderStyle`、`ValidationRange`、`MappingValidationBenchmarks.RegexCacheHit`、`Program.ResourceProbe.RunScenario`、Round 10 benchmark/resource 产物
- 问题：BenchmarkDotNet 文件已经生成，但多个方法展示的输入参数没有进入实际 workload；Header style benchmark 未使用 HeaderAttribute；validation benchmark 未覆盖 ValidationRangeIndex 规模；资源探针 retained 指标未保持 UniqueTracker workload 存活。
- 证据：failure/validation fixture 固定为单行；HeaderStyle 对 10K/100K 参数仍截断为 1000 行；`BenchmarkRow` 无 HeaderAttribute；三组报告的 allocated 基本不随标签变化；`GC.KeepAlive` 仅覆盖 `plans`，不覆盖 tracker/values；没有 destination Capacity/retained 指标。
- 影响：当前 JSON/Markdown 不能支持 failure workbook、HeaderAttribute style、validation index 在声明输入规模下的绝对性能、分配和 retained-capacity 结论，可能误导发布资源评估。
- 修复目标：让每个 benchmark 的参数只描述该方法实际消费的 workload，并补齐 ValidationRangeIndex、HeaderAttribute style 和 retained-capacity 的可审计证据。
- 明确修复要求：
	1. 为 failure workbook 创建随失败行数增长的 fixture，或移除无效 `RowCount` 并使用真实的 `FailureRowCount` 参数。
	2. 为 validation 建立直接的大范围/重叠/多规则 `ValidationRangeIndex` workload，输入规模必须真实影响查询次数或规则数量。
	3. style benchmark 使用带 `HeaderAttribute` 的宽模型或等价真实生产路径；不要把 1000 行 workload 标成 10K/100K。
	4. 将 Mapping benchmark 的无关参数按 benchmark 类型拆分，避免 Regex/plan/unique 结果携带未消费参数。
	5. 在资源探针强制 GC 后测 retained 时，对计划和 UniqueTracker/values workload 分别建立明确存活/释放阶段并使用 `GC.KeepAlive`；如要声称 stream retained capacity，记录 fresh destination 的 Capacity/峰值或提供等价独立指标。
	6. 重新运行受影响分组，保留独立 JSON/Markdown，并在执行报告中只报告实际输入规模和绝对指标；无历史 baseline 时不得给出相对提升结论。
- 修复后的验证方式：Reviewer 对照源码确认每个参数进入 workload；读取 JSON/Markdown 核对输入规模、Mean、Allocated、GC、环境；检查 failure/style/validation 指标随真实规模变化；核对资源 JSON 中各 workload 的 sampled peak、retained、工作集和存活阶段定义。

## 7. 独立验证结果

| 验证 | 结果 |
| --- | --- |
| Round 10 Header/Regex 专项 net8 | PASS，4/4 |
| Round 10 Header/Regex 专项 net6 | PASS，4/4 |
| 迁移 fence/当前 major consumer 专项 | PASS，2/2 |
| net8 Unit 全量 Release | PASS，329/329 |
| net6 Unit 全量 Release | PASS，329/329 |
| net8 Integration Release | PASS，12/12 |
| net6 Integration Release | PASS，12/12 |
| Docs consumer/fence net8 | PASS，9/9；迁移文档已实际覆盖 |
| `dotnet build Bing.Offices.sln -c Release --no-restore` | PASS，0 errors；既有弃用/旧 TFM 警告 |
| `dotnet pack Bing.Offices.sln -c Release --no-build --no-restore` | PASS |
| changed C# diagnostics | PASS，无错误 |
| `git diff --check` | PASS，无空白错误；有行尾转换提示 |
| GenerationId/manifest SHA/source SHA | PASS，三份证据代际一致，四个 source SHA 全部匹配 |
| Excel COM roundtrip | BLOCKED，四次 `Workbooks.Open` 均失败 |
| NPOI roundtrip 二次解析 | BLOCKED，四个 roundtrip 文件均缺失 |
| LibreOffice/WPS | BLOCKED，当前环境未发现客户端 |
| BenchmarkDotNet 文件格式/环境字段 | PASS，JSON/Markdown、Mean、Allocated、GC、CPU、SDK/runtime 存在 |
| Benchmark workload 与参数一致性 | FAIL，failure/validation/style/regex 存在未消费或失真参数 |
| Resource probe | PARTIAL，16/16 exit 0；真实类型 workload 已使用，但 retained 存活定义不完整 |

## 8. 计划逐项复审结论

| 计划项 | 复审结果 | 说明 |
| --- | --- | --- |
| RH-101/RH-103/RH-104 | PASS（既定 FIX 合同） | 前序 P0 正确性合同未发现回归。 |
| RH-401 | PASS | 有界 Regex 缓存、timeout、FIFO、并发和文档继续有效。 |
| RH-402 | PARTIAL | UniqueTracker 功能成立；ValidationRangeIndex 的直接性能/分配证据仍缺失。 |
| RH-403 | PARTIAL | HeaderAttribute 双格式功能测试通过；对应 performance benchmark 未走 HeaderAttribute。 |
| RH-404 | PARTIAL/FAIL | 原始产物已生成、人工 payload 已移除，但 workload 参数和 retained 指标仍不可信。 |
| RH-201/RH-202 | PARTIAL/BLOCKED | 当前 major 兼容证据有效；breaking table/去全局化等待批准。 |
| RH-502 | PASS（当前 major 文档范围） | 迁移示例已编译执行，文档与未批准状态一致。 |
| RH-601 | PASS（本机门禁） | Build、pack、net6/net8 Unit/Integration、Docs consumer 均通过。 |
| RH-602 | PARTIAL/BLOCKED | 四个 source 与哈希证据可靠，真实客户端 roundtrip 仍为空。 |
| RH-603 | PASS | 准确保持 NO-RELEASE，未执行发布。 |

## 9. 最终 Checklist

- [x] FIX-001 至 FIX-003 既有正确性合同保持关闭。
- [x] FIX-005 Regex 缓存合同保持关闭。
- [x] FIX-007 迁移示例和 fence 覆盖已关闭；未批准 breaking change 保持延期。
- [x] HeaderAttribute XLSX/XLS 双格式功能测试通过。
- [x] BenchmarkDotNet JSON/Markdown 和真实类型资源探针产物已生成。
- [ ] FIX-004 完成真实办公客户端 roundtrip、截图、roundtrip SHA 和 NPOI 结构 PASS。
- [ ] FIX-006 让 failure/style/validation/regex 参数与实际 workload 一致，并补齐 retained-capacity 证据。
- [x] net6/net8 Unit 与 Integration、Docs consumer、Release build 和 pack 通过。
- [ ] `execution.md` 历史中段重复机器元数据仍存在（LOW，不新增本轮 FIX）。
- [x] 未自动 commit、push、tag 或 publish。
- [x] 保持 NO-RELEASE。

## 10. 最终状态

`NEEDS_FIX`

Reviewer 未实施修复。下一轮应继续处理 `FIX-004`，并按 recommended scope 修复 `FIX-006`；`FIX-005`、`FIX-007` 不应重复开启。
