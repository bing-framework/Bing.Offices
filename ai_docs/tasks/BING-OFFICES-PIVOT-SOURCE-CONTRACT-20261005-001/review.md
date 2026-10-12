<!-- AI_REVIEW_STATUS: PASS -->
AI_TASK_ID: BING-OFFICES-PIVOT-SOURCE-CONTRACT-20261005-001
AI_REVIEWED_AT: 2026-10-05T04:41:28.6183063+08:00

# 独立代码审查

## 审查结论

PASS。未发现未解决的 MUST_FIX、SHOULD_FIX 或 OPTIONAL 问题。

本次审查对照 plan、COMPLETED execution、当前源码、测试、文档、Git Diff、API 产物和验证汇总进行。历史 review-progress-20261005.md 与 review-todo-20261005.md 中的编译、编码、职责测试和交付缺口均已有对应修复及新鲜证据；历史失败仍按原记录保留，没有用旧通过结果替代当前候选。

## 计划验收矩阵

| 计划项 | 结果 | 独立复核依据 |
| --- | --- | --- |
| SOURCE-001 来源预检与诊断 | PASS | `ClosedXmlReportWriter.ApplyPivotTables` 在创建目标 Sheet 和 PivotTable 前调用 `ClosedXmlPivotSourceValidator.Validate`；校验器先扫描完整来源区域的公式和 Excel 错误，再读取表头、分组与求和值。异常使用实际 `IXLCell` 的一基行列、来源 Sheet 和透视名称。 |
| SOURCE-002 职责与组合测试 | PASS | 双 TFM 最新来源职责各 66/66；覆盖公式位置及缓存状态、来源外公式、完整缓存、类型化分组、空值、标题边界、各字段错误、后序透视失败、流与文件保护、取消、临时文件清理和单次异常观察。 |
| SOURCE-003 文档、API 与交付 | PASS | 模型 remarks、透视与迁移文档、符号测试映射、审批和 execution 完整；八个生产包成功，隔离 net6/net8 消费者均输出 `package-consumer-ok`；双 TFM 公共成员差异均为空。 |

## 真实调用链

`ClosedXmlExcelExporter.ExportWorkbookCore` 先生成请求中的业务 Sheet，再进入 `ClosedXmlReportWriter.ApplyPivotTables`。每个透视定义先解析来源区域并执行 `ClosedXmlPivotSourceValidator.Validate`；公式通过 `HasFormula` 拒绝，Excel 错误通过缓存类型拒绝。只有全范围公式/错误扫描结束后才调用 `GetString`、读取分组 `Value`、判断求和值类型并创建目标 Sheet、PivotTable 和缓存。工作簿 `SaveAs` 位于全部透视处理之后，因此预检失败不会序列化到调用方流。

同步流入口在 `SaveAs` 前失败；异步流入口先在内部缓冲完成工作簿，成功后才异步复制到调用方流。文件入口复用默认原子提交器。职责测试同时证明首个透视已处理、后一个来源失败时，调用方流、既有目标文件和临时文件状态仍符合合同，异常只观察一次。

来源公式覆盖表头、分组、求和与未选缓存字段，支持函数、未知函数及新鲜、过期、缺失缓存都执行相同拒绝规则。来源区域外公式重新打开后仍保留。来源错误、完整空行、非数值金额、全空金额、空白及重复表头均返回结构化来源坐标；空分组和部分空金额允许。字段比较使用 `OrdinalIgnoreCase` 且不 trim，带空格精确匹配和未带空格拒绝均有直接测试。展开预算使用 `long` 算术，并覆盖精确边界与越界来源坐标。

## API 与身份

`artifacts/api-before/api-member-diff.json` 和 `artifacts/api-verified/api-diff.json` 的 net6.0、net8.0 均为 `{}`。根 `build/api-snapshot-baseline.json` 与已完整验证的 `api-approved-candidate.json` SHA-256 相同；审批仍为 `user:pivot-table-mvp-2026-10-04` / `2026-10-04`。本阶段新增类型为 ClosedXML Provider 内部类型，没有新增、删除或修改公共成员，也没有新增生产友元、依赖或公式计算选项。

## 验证证据

- 来源职责初次最终结果：net6/net8 各 50/50。
- 追加边界后的来源职责：net6/net8 各 66/66，0 失败、0 跳过。
- Provider 合同：net6/net8 各 441/441。
- ClosedXML 回归：net6/net8 各 339/339。
- `Category!=Large` 常规门禁：27 份 TRX，共 4,473/4,473，0 失败、0 跳过，命令退出码 0。
- 身份刷新后的 API 合同：net6/net8 各 11/11；最新文档职责 net8 为 19/19。
- 解决方案 Release 构建退出码 0，0 错误；八个 2.0.0 生产包顺序打包成功。
- 隔离 NuGet 缓存消费者：.NET 6.0.36 与 .NET 8.0.31 均退出码 0，并真实执行公式失败坐标和流保护合同。
- `git diff --check` 无空白错误；本阶段 C# 文件符合 UTF-8 BOM、LF 和末尾换行约定。

常规门禁完成后只追加职责测试、文档和已验证的 API 身份元数据，生产候选未改变；新增部分随后已通过双 TFM 来源职责、API 合同和文档职责验证，因此没有把未重新执行的用例计入 4,473 项结果。

## 风险边界

本阶段明确收紧来源输入：公式结果必须先由应用或显式计算接口计算，再以字面值写入；数字文本、日期和布尔值不能隐式参与 Sum。金额仍遵循 Excel 双精度存储，精度与舍入由应用负责。

客户端首次显示与刷新、Excel/WPS/LibreOffice 兼容、多透视缓存复用和资源预算、高级透视、NPOI 透视缓存及模板已有透视保留属于后续计划，不构成本任务未完成项。当前 OOXML 与缓存结构通过不应解释为这些后续能力已经完成。
