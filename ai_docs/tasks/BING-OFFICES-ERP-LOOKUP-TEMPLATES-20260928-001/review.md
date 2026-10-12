<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-ERP-LOOKUP-TEMPLATES-20260928-001
AI_REVIEWED_AT: 2026-09-29T01:09:44Z

# Round 2：计划差距与 TODO

## 结论

主要公共入口和业务路径已经实现，尚不能宣布完整验收通过。本次确认 2 项实现问题、1 组专项测试缺口，共 3 组 OPEN_ACTIONABLE；另有最终包验证待刷新。原定后续 4 个阶段不属于本轮漏实现。数量按问题组统计，不代表工作量百分比。

本次只审查，未修改业务源码、正式测试、plan.md 或 execution.md。上一轮独立报告保存在 review-round1.md。执行记录中“全部 COMPLETED”不替代当前真实行为验证。

## 计划验收矩阵

| 计划项 | 状态 | 本轮复核 |
|---|---|---|
| P1 公共契约、旧重载、快照与互斥 | PASS | Workbook/Entity LookupList、ListSourceName、DataValidation、四参数导入选项均存在并接入 |
| 隐藏字典表、文本、名称、模板冲突 | PASS | 两 Provider 结构测试、模板冲突与辅助表顺延证据存在 |
| 手工名称前向引用与一维来源预检 | PASS | Round 1 已调整建立名称顺序并补充直接测试 |
| NPOI 内联/名称并存及旧内联兼容 | FAIL | XLSX XML 定位修复存在，但 HSSF 内联导入发生回归，见 FIX-010 |
| 名称解析、局部作用域及缓存 | PARTIAL | 特殊 Sheet 名、局部优先、HashSet 缓存已实现；NPOI 导入取消令牌未传入，见 FIX-011 |
| Entity 原生校验顺序及原始坐标 | PASS | ClosedXML 整行预校验、NPOI 源行映射及分页小计坐标测试存在 |
| MiniExcel/SpreadCheetah 提前拒绝 | PASS | 仅字典/仅名称规则的同步异步、流和文件直接测试已补齐 |
| 专项资源与格式测试 | PARTIAL | 预取消及提交失败已覆盖；成功异步、字典处理中途取消及部分 Unsupported 用例仍欠缺 |
| API、文档、包消费者 | PARTIAL | 已有双 TFM +14/-0、文档与 Consumer 证据；修复后使用新包重新验证消费者 |

## 验证证据

- 复核当前源码、Git 改动、任务记录及职责测试；复用执行记录中最终 Release 构建和 Category!=Large 全量成功结果，未无故重跑全量。
- 本次独立 net8.0 探针位于 artifacts/review-erp-lookup-round2，引用当前 NPOI 测试 Release 输出（不是旧 Consumer 包）。生成真实 XLS/XLSX 内联 Yes/No 校验并通过公开 Import 导入合法 Yes。
- 探针成功运行，退出码 0：XLS 的 Formula1 为 null、ExplicitListValues 为 Yes/No，结果 0 行和 WorkbookValidation“不支持的规则”；XLSX 结果 1 行且无错误。
- 取消问题由真实调用链确认，未把它描述为已经运行的中途取消测试。

## FIX-010：恢复 HSSF 显式内联列表导入

- 严重程度：HIGH；处理要求：MUST_FIX；状态：OPEN_ACTIONABLE。
- 问题：TryGetExplicitListValues 在 Formula1 为空时提前返回，尚未读取有效的 ExplicitListValues。
- 证据：src/Bing.Offices.Npoi/Bing/Offices/Imports/NpoiWorkbookValidationPipeline.cs:196；上述真实工作簿探针已复现。
- 影响：合法 XLS 内联下拉值被拒绝；共享管线也影响使用该规则的 Entity 导入，与“显式列表保留原行为”冲突。
- 修复目标：同时支持 HSSF 无 Formula1 的显式列表与名称来源，不重新引入名称误判或 XML 混写。
- 修复要求：调整解析次序；补齐 Workbook 和 Entity 的显式列表合法/非法直接回归。
- 验证方式：XLS/HSSF、XLSX/XSSF × net6.0/net8.0；断言接纳行、完整错误及坐标，并保留混合列表回归。

## FIX-011：将导入取消令牌传入名称字典扫描

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题：TryGetNamedListValues 虽接收可选 token，但导入调用使用默认值，逐行取消检查无法响应调用方取消。
- 证据：同文件 Validate:37、ValidateCell:78 不接收 token；TryGetExplicitListValues:195 调用 TryGetNamedListValues 时未传 token；后者:315-316 默认 default。只有导出预检 IsSupportedNamedListSource:337 明确传入 token。
- 影响：Workbook/Entity 导入首次扫描大名称区域时不能及时取消。execution.md 对 FIX-006“扫描期间检查取消”的完成描述不充分。
- 修复目标：从 Workbook/Entity 导入入口贯通原 token 到名称解析与扫描。
- 修复要求：保持原取消异常、流所有权、错误观察及缓存隔离合同；不得用 Task.Run 包装。
- 验证方式：确定性的扫描中取消测试，避免依赖毫秒计时或无关大规模性能矩阵；覆盖双格式和两个入口。

## FIX-012：补齐新增字典路径的专项测试

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；状态：OPEN_ACTIONABLE。
- 问题与证据：两个 Provider 的 LookupListExportStructureTest 成功结构用例均同步；新增异步取消用例只覆盖 PreCanceled。通用旧路径的中途取消不能直接证明新增字典循环的行为。NPOI NamedList_UnsupportedDefinitions_ShouldReturnStructuredError 未提供联合区域和外部引用数据，尽管注释声称包含外部引用。
- 影响：计划中的同异步和新路径资源合同尚缺直接证据；不是断言所有相应实现已失败。
- 修复目标：补齐有代表性的职责矩阵，不重复无关资源矩阵。
- 修复要求：添加成功异步字典输出并完整读回；添加字典处理中途取消，断言文件目标保持、流所有权及临时文件；NPOI 添加联合/外部引用 Unsupported 用例及完整错误断言，修正测试注释。
- 验证方式：NPOI XLS/XLSX、ClosedXML XLSX，受影响测试双 TFM；共享实现可复用既有证据，但须记录覆盖关系。

## 收口验证待办（不单列功能缺陷）

修复后按影响范围先跑职责测试，再跑受影响集成；生成包含最终修复的 NuGet 包并运行 net6.0/net8.0 Consumer。Round 1 沿用了修复前包消费者证据，它可证明包入口曾运行，不能代替最终修复版本验证。API 预期仍保持已审批 +14/-0，无需为内部修复新增基线差异。更新执行记录与符号映射，完成差异、编码检查；不重跑无关大性能矩阵。

## 后续独立阶段（DEFERRED）

1. 填写区域保护：设计工作表保护与可编辑区域公共契约。
2. 级联字典：设计依赖关系、父子值以及空值/失效值行为。
3. 分批导入扩展：设计多 Sheet、动态列、关系与失败工作簿的组合边界。
4. 分组分页组合与透视报表：分别设计布局组合和公共报表模型。

建议顺序：FIX-010 → FIX-011 → FIX-012 → 最终包与验收记录 → 后续阶段 1～4。本次审查到此停止，不自动进入修复。
