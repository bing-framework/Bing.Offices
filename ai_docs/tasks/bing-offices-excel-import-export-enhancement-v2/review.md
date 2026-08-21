<!-- AI_REVIEW_STATUS: PASS_WITH_ISSUES -->
AI_TASK_ID: bing-offices-excel-import-export-enhancement-v2
AI_REVIEWED_AT: 2026-08-21T14:09:05.1416944+08:00

# Review Fix Round 3 独立复审报告

## 验收摘要

最终结论：`PASS_WITH_ISSUES`。

本次以当前 `plan.md`、当前已更新的 `execution.md`、当前源码、测试和 Git 工作树为证据，优先逐项复核上一轮 `FIX-001` 至 `FIX-006`，重点验证 Round 3 的 `FIX-004`、`FIX-005`。两项此前未闭合的 HIGH/MUST_FIX 均已真正接入生产路径并有直接回归证据；未发现新的 BLOCKER/HIGH 或与本轮修复有关的回归，因此不生成新的 `FIX-xxx`。

本 Review 未修改业务代码、测试代码、`plan.md` 或 `execution.md`；仅更新本文件。

## Review 边界与 Git 分析

- 计划涉及 Abstractions/Core/NPOI 的 Workbook API、导入导出管线、失败工作簿、Workbook Validation、图片、公共面、Docs Consumer 和测试。
- 当前工作树为 `master...origin/master [ahead 1]`，包含大规模既有删除、新增和重构；其归属不能仅按 Round 3 拆分。本轮直接关联的改动集中于 `ValidationRangeIndex`、`ExcelColumnPlan`、NPOI Importer/Exporter 和两份 Excel 回归测试，与 `execution.md` 的 Round 3 记录一致。
- 未发现本轮 Diff 引入新的公开 NPOI API、旧 `ExcelCompatibilityTestAdapter` 或遗留的 `ConvertValue`/`ConvertDynamicValue`/`ValidateDynamicValue` 并行转换校验方法。
- `git diff --check` 通过；仅输出既有 CRLF/LF 规范化提示，无空白错误。

## 上一轮 FIX 逐项复核

| FIX | 结论 | 当前证据 |
| --- | --- | --- |
| FIX-001 | `RESOLVED` | `ExcelImportErrorCollector.Add()` 写入第 N 条错误时立即标记截断；根 collector 仍在结构、行和关系路径共享。`Import_StructureErrorAtMaxErrors_ShouldMarkTruncatedAndStopLaterSheets` 保持通过。 |
| FIX-002 | `RESOLVED` | `NpoiExcelImporter`、`NpoiExcelExporter`、`ColorResolver` 均为 internal；NPOI 程序集只公开 `ExcelNpoiServiceCollectionExtensions.AddNpoi()`。当前 Public API 快照与 Docs Consumer 均通过。 |
| FIX-003 | `RESOLVED` | `ErrorRowsOnly` 继续创建独立 Workbook，复制失败 Sheet/行、公式、样式、Merge、Validation、图片和来源列。`Import_ErrorRowsOnly_ShouldCopyAndReorderOriginalFailureRows` 仍为直接覆盖。 |
| FIX-004 | `RESOLVED` | `ValidationRangeIndex` 现为行区间树加节点内列区间树。查询仅沿目标行路径并按目标列缩小候选；同一 validation 通过引用去重，矩形未按 Cell 展开。`ValidationRangeIndex_OverlappingRows_ShouldLimitColumnCandidates` 以 200 个同一行但不相交列范围直接断言候选检查数小于全部规则。 |
| FIX-005 | `RESOLVED` | `ExcelColumnPlan` 预绑定 converter、attribute/named rules，并统一执行 `ConvertFrom`、`ConvertTo`、ValueMap、默认类型转换和 `WriteValue`。Importer 的 fixed/typed dynamic 均消费计划转换和验证，Exporter 均消费计划转换与写入；动态字典读取/写入仅保留数据载体适配。固定、动态 converter 绑定计数及既有 ValueMap/DataType 回归均通过。 |
| FIX-006 | `RESOLVED` | 旧兼容 Adapter 未检出；泛型关系 Key/comparer 与真实来源坐标回归继续存在，完整公开成员快照和 Docs Consumer 测试通过。 |

## 计划验收矩阵

| 计划项 | 状态 | 独立证据 |
| --- | --- | --- |
| P1-002 统一不可变 Column Plan | `PASS` | 每 Sheet 创建期间绑定 converter/rule；逐 Cell 不再查询 converter/命名规则。`ExcelColumnPlan` 承担双向转换、ValueMap 和写入，Reader/Writer 只对 fixed/dynamic 的值容器适配分支。专项 8 项、Unit net6/net8 全量通过。 |
| P4-001 Unique、Ignore、关系 Key | `PASS` | `DuplicationExcelValidationRule` 经计划的 attribute binding 统一进入转换后校验；失败行继续回滚当前行 duplicate 变更。关系泛型 Key/comparer 直接测试仍通过。 |
| P5-001/P5-003 错误上限与 ErrorRowsOnly | `PASS` | 根级有界错误 collector、独立失败工作簿及其结构复制均保留，结构上限与 ErrorRowsOnly 回归仍在全量 Unit 中通过。 |
| P6-001 Workbook Validation Range Index | `PASS` | 预编译范围保持矩形，不进行 Cell 展开；二维候选索引解决同一行、无关列范围的线性扫描问题。大矩形、离散范围和重叠行专项均通过。 |
| P8-001/P8-002 公共面与消费者 | `PASS` | NPOI 可见性保持收敛，Public API 精确成员测试由全量 Unit 覆盖；Docs Consumer net8 `2/2` 通过。 |

## 功能、契约与维护性 Review

- `ValidationRangeIndex` 的列树在左/右查询时按已排序的起止边界提前停止，目标列不相关的范围不会进入 `AddIfMatching`；对真实同时覆盖目标行、列的规则仍需检查，这是结果正确性所需而非无关规则扫描。
- `ExcelColumnPlan.ConvertFrom()`/`ConvertTo()` 先使用预绑定 converter，再处理 ValueMap 和类型回退；`WriteValue()` 统一维持 formatter/decimal-scale 行为。全量 Unit 独立复现并通过了此前统一改造曾暴露的动态 `DataType` 数值 Cell 与 fixed ValueMap 回写风险。
- Importer 的 raw Required/Regex 与 converted attribute/named validation 是有意保留的校验阶段差异，但两阶段均遍历 Column Plan 的预绑定规则，动态列不再被跳过。图片类型转换共用 `ConvertImages()`，差异仅为 dynamic dictionary 或固定 Setter 的目标写入方式。
- 并发请求隔离已有 `StreamPipeline_ConcurrentRequestOptions_ShouldRemainIsolated`，本次专项复测通过；当前计划对象在 Sheet 执行期间构建，未发现请求级绑定进入静态 TypeMap 缓存。

## 性能与资源 Review

- P6-001：范围索引构建按 validation 矩形保存，空间复杂度与范围数量相关，不随覆盖 Cell 数量展开。候选检查白盒断言验证了相同行、不同列的关键退化反例。
- P1-002：converter `CanConvert`、命名规则解析和 attribute rule 绑定均在 `CreateColumns()` 阶段完成；逐 Cell 路径只迭代已绑定数组，未检出 LINQ 服务集合查找或规则名查找。
- 未发现 Round 3 对根级 `MaxErrors`、图片资源上限、流所有权或取消检查路径的改动。

## 验证结果

- `dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore --filter ...`：通过，`8/8`。
- `dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release -f net8.0 --no-restore`：通过，`135/135`。
- `dotnet test .\tests\Bing.Offices.Tests\Bing.Offices.Tests.csproj -c Release -f net6.0 --no-restore`：通过，`135/135`。
- `dotnet test .\tests\Bing.Offices.Tests.Integration\Bing.Offices.Tests.Integration.csproj -c Release -f net8.0 --no-restore`：通过，`10/10`。
- `dotnet test .\tests\Bing.Offices.Docs.Tests\Bing.Offices.Docs.Tests.csproj -c Release -f net8.0 --no-restore`：通过，`2/2`。
- 改动相关源码和测试的编辑器诊断：无错误。
- `git diff --check`：通过，仅 CRLF/LF 提示。

上述构建仍有既有 `netcoreapp3.1` package-support、obsolete `ICellValueConverter`、XML 注释和一项 xUnit analyzer 警告；没有测试或编译失败。

## 问题与残余风险

- `LOW`：`dotnet format --verify-no-changes --no-restore` 仍受仓库既有 `FINALNEWLINE`、`IMPORTS`、`IDE1006` 基线影响。本轮未运行会写入文件的格式化命令；该项不阻塞本次 Review Fix 验收。
- `NOT_VERIFIABLE`：当前环境未提供真实 Excel/LibreOffice 互操作验证条件，无法独立确认高级 OOXML 部件在外部 Office 产品中的保真度。NPOI 重开回归与集成测试已覆盖可自动验证部分。
- net6/net8 已独立验证；net5 运行时不在本次环境验证范围内。

## 最终验收 Checklist

- [x] 已读取当前 `plan.md`、当前 `execution.md`、旧 `review.md`、当前源码和 Git 工作树。
- [x] 已逐项复核上一轮 FIX-001 至 FIX-006。
- [x] `FIX-004` 的行列双维候选索引和性能回归测试已独立验证。
- [x] `FIX-005` 的共享执行计划、固定/动态 converter 预绑定、ValueMap/DataType 和并发隔离已独立验证。
- [x] Unit net8/net6、Integration net8、Docs Consumer net8 通过。
- [x] 无未解决的 BLOCKER/HIGH/MUST_FIX。
- [x] 未修改业务代码、测试、计划或执行报告。

## 结论

Round 3 的两个 MUST_FIX 已闭合，先前其余 FIX 未回归。结论为 `PASS_WITH_ISSUES`，原因仅为不阻塞的格式基线和外部 Office/LibreOffice 互操作不可验证；无需进入 `review-fixer`。

## 当前复核

本次再次读取当前 `execution.md` 的 Round 3 记录、当前 Git 工作树及 `ValidationRangeIndex`、`ExcelColumnPlan` 实现。专项 Unit net8 复验 `6/6` 通过，覆盖 `FIX-004` 的行列候选缩减，以及 `FIX-005` 的固定/动态 converter 预绑定、动态 DataType、ValueMap 回归和并发请求隔离。当前无新增或回归的 `FIX-xxx`，结论维持 `PASS_WITH_ISSUES`。
