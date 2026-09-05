<!-- AI_REVIEW_STATUS: PASS_WITH_ISSUES -->
AI_TASK_ID: BING-OFFICES-RC-HARDENING-20260904-001
AI_REVIEWED_AT: 2026-09-04T09:35:00.000Z
AI_REVIEW_ROUND: 2

# 独立 Review Round 2

## 结论

代码级 `PASS_WITH_ISSUES`：Round 1 的 `FIX-001` 已完成，未发现新的 P0/P1 代码缺陷；发布级门禁仍为 `BLOCKED / No-Go`。

Round 1 原始证据保留在 `review.md`，其 `NEEDS_FIX` 状态不被覆盖。本文记录修复后的独立复核结论。

## FIX-001 复核

- 状态：`CLOSED`
- `ExcelResourceLimits.Validate()` 仅在 `MaxZipCompressionRatio.HasValue` 时检查非正数、NaN 和 Infinity。
- `MaxZipCompressionRatio = null` 的直接 Unit 测试通过。
- `0`、负数、NaN、正无穷、负无穷的拒绝测试通过。
- net8/net6/netcoreapp3.1 的 ZIP 专项均为 `16/16`。
- `get_errors` 对相关源码和测试目录为 `No errors found`。

## 安全与边界复核

- `NpoiXlsxZipPreflight.Validate` 在 `WorkbookFactory.Create` 前调用。
- ZIP entry 数、单 entry/总解压大小、压缩比、sharedStrings、styles、worksheet、重复 entry、异常路径均有预算或结构检查。
- XML reader 禁止 DTD，`XmlResolver = null`，并设置文档字符上限；未宣称独立 XML 最大深度硬上限。
- 取消令牌在预检入口、entry 循环和 XML reader 循环中检查。
- XLS/OLE 没有被误写成具备 XLSX ZIP 等价的内部预检。
- 生产程序集之间无 IVT；日期源码未发现宽松 `DateTime.Parse`、`DateTimeOffset.Parse` 或本地时区依赖。
- consumer 的最终 `project.assets.json` 无 `projectReferences`，`libraries` 中无 project 类型，net6/net8 均以最终本地 nupkg 成功运行。

## 剩余发布问题

1. formal API candidate hash 尚未得到维护者批准，不能更新正式 baseline。
2. 性能/资源预算未批准，且完整 Failure Workbook 双 DOM、100k/1M 和全矩阵证据未完成。
3. DataTable 显式兼容类、Office exception 家族、UniqueTracker 和剩余 public execution detail 的治理仍未完成。
4. 同版本本地 nupkg 重发和隔离缓存成功不等于 feed、clean clone 或正式发布证明。
5. 导出元数据仍使用 `DateTime.Now`；这是确定性增强项的后续风险，不是本轮输入日期 parser 的回归。

## Review 裁决

- P0/P1 代码修复项：`0 open`。
- 发布治理/证据阻断：保留，任务整体 `BLOCKED / No-Go`。
- Round 1 `review.md` 不修改，作为独立 Reviewer 原始证据保留。
