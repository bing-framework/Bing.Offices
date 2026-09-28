<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-ENTITY-FOOTER-SUM-20260928-001
AI_EXECUTION_FINISHED_AT: 2026-09-28T08:48:55Z

# 实施执行报告

## 执行结论

新增 `ExcelEntityListFooterBuilder<TItem>.FormulaSumContiguousRowsAbove`，为本尾部上方连续的明细生成原生 Excel 求和公式。公式根据明细快照数量、构建时冻结的 `GapRows` 和公式单元格相对行定位；空明细写入 `=0`。最终 Footer 若跨越分页或分组小计，布局构建阶段拒绝该便捷方法；中间小计自身可使用。现有 `Cell`、`Formula`、Provider SPI 接口及构造签名不变。

## 任务信息

- 计划：`ai_docs/tasks/BING-OFFICES-ENTITY-FOOTER-SUM-20260928-001/plan.md`
- 前序：`ai_docs/tasks/BING-OFFICES-UTOPA-BUSINESS-LAYOUT-20260927-001/continuation-16-execution.md`
- 范围：仅本仓库；Utopa 外部样本保持只读。

## 计划执行情况

| 项目 | 结果 |
| --- | --- |
| 公共布局契约与快照隔离 | 完成；`GapRows` 的最终值在构建时冻结，调用顺序不影响生成公式 |
| 连续范围与冲突校验 | 完成；分页/分组小计可用，跨中间小计的最终 Footer 在布局阶段拒绝 |
| NPOI HSSF/XSSF、ClosedXML XLSX | 完成；复用现有公式值 SPI，无 Provider 友元或实现分叉 |
| 文档、包消费者、API 审批与追溯 | 完成；新成员已登记到原任务的审批与生产符号映射 |

## 修改文件

生产：`ExcelEntityListExtensions.cs`、`ExcelEntityLayout.cs`。测试：NPOI、ClosedXML Provider 测试，Docs 测试及第三方包消费者。文档：`docs/excel/dynamic-columns.md`、本计划与执行记录、原任务 API 审批和生产符号映射、双 TFM API baseline。

## API/数据/配置变化

双 TFM 各新增一个公共方法，无公共成员删除、签名修改、包版本修改或数据迁移。来源列接受 A 到 XFD 的纯列字母；具体 Provider 的物理列边界仍由导出工作簿检查。公式不由本库求值。无新增配置项。

## 测试结果

- 新能力定向：NPOI net6.0/net8.0 各 13/13；ClosedXML net6.0/net8.0 各 8/8。覆盖空/非空、GapRows 前后配置、构建后隔离、公式相对下一行、分页和分组小计、非法列、冲突、导入边界及目标流保护。
- 受影响 Provider 常规 `Category!=Large`：NPOI 双 TFM 各 709/709；ClosedXML 双 TFM 各 213/213。
- Docs 原文示例：net8.0 10/10，19 个 C# 围栏可编译执行。
- 解决方案常规测试：更新前只有旧 API 快照断言失败，其余项目均通过；基线更新后完整 API 集成测试 net6.0/net8.0 各 32/32 通过。
- 第三方消费者从新打的本地包隔离还原并运行，net6.0/net8.0 均输出 `third-party-public-only-provider-ok`。

## Build/Typecheck/Lint/Format

`Bing.Offices.sln` Release 构建 0 错误；八个生产 NuGet 包本地打包成功。API 双 TFM 差异仅为计划批准的方法，`removed=[]`，候选身份与基线比较通过。严格 UTF-8/BOM/EOL 与 `git diff --check` 已检查；保留存量 API baseline 的 CRLF/无 BOM/无末尾换行格式例外。

## 计划偏差

隔离还原首次从本机全局缓存选中同版本旧 Abstractions 包，导致消费者缺少 `Entities`。改用临时精确包源映射并以全新缓存还原后通过；正式仓库 NuGet 配置未修改。任务状态脚本首次写入 `.agents/runtime` 遭沙箱拒绝，按项目规则完成授权写入后登记成功。

## 基线问题与已知问题

net6.0 的 `Microsoft.Bcl.Memory` 支持警告及分析器离线还原警告为现有构建环境问题，未改变依赖。`INDEX` 求和便捷方法仅适用于连续明细；跨中间小计的最终总计继续使用聚合 `Cell` 或业务显式 `Formula`。

## 风险与回归关注点

关注第三方 Provider 对现有 `ExcelEntityFooterFormulaValue` 的写入支持，以及 XLS 物理列边界。公共 Provider 合同和官方 Provider 测试已覆盖本次路径；未运行无关的大规模性能矩阵。

## Reviewer 注意事项

新方法不改变现有 `Formula` 或 `Cell` 的行为。新增代码在 Build 时捕获 `GapRows`，防止外部保留构建器引用并事后修改布局。`FormulaSumContiguousRowsAbove` 对最终 Footer 的冲突校验不限制分页/分组小计本身。

## Git 状态

工作树保留此前未提交的阶段改动；本阶段未自动 git add、commit、push、创建 PR 或发布。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
RELEASE_STATUS: NOT_EVALUATED
OPEN_ACTIONABLE: 0
