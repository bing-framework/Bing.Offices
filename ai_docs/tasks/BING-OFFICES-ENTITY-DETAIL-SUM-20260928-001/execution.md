<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-ENTITY-DETAIL-SUM-20260928-001
AI_EXECUTION_FINISHED_AT: 2026-09-28T10:28:36.000Z

# 实施执行报告

## 执行结论

计划内的最终 Footer 跨小计明细求和能力已完成。NPOI XLS/XLSX、ClosedXML XLSX 均按实际明细行段生成原生 SUM 公式；分页或分组小计、间隔行和签字行不参与求和，空明细生成 `=0`。双 TFM 的职责级测试、共享 Provider 合同、第三方 NuGet 消费者和 API 门禁均通过。

## 任务信息

- 计划：[plan.md](plan.md)。本阶段只修改 Bing.Offices，不修改外部 Utopa 样本。
- 基于现有 Entity Layout Footer 与 Provider SPI 扩展；全部旧公开构造、接口和 Fluent 调用保持兼容。

## 计划执行情况

1. 新增最终 Footer 的 `FormulaSumDetailRowsAbove` 和隐藏于 IntelliSense 的公开 Provider SPI 值对象；验证来源列、明细闭区间、SUM 参数分段及公式长度。
2. NPOI、ClosedXML 收集真实写出的分页/分组明细行段，最终 Footer 写入原生公式；中间小计使用该方法在布局构建阶段拒绝。
3. 补充 HSSF/XSSF/XLSX 业务场景、共享 Provider 合同、NuGet 消费者、中文 XML、文档、成员审批、双 TFM API 快照与生产符号追溯。

## 已完成事项

- 普通明细、分页小计和分组小计路径共用最终 Footer 公式语义；已有 `Formula`、`FormulaSumContiguousRowsAbove`、marker 导入边界与模板/流行为保留。
- 真实工作簿测试覆盖空/非空明细、多行小计、`GapRows`、数字格式、公式文本、导入截断和非法中间小计配置。
- [dynamic-columns.md](../../../docs/excel/dynamic-columns.md) 和 [entity-provider-contracts.md](../../../docs/excel/entity-provider-contracts.md) 说明两类求和方法的适用范围。

## 部分/未完成事项

无计划内遗留。未执行与本变更无关的大型性能矩阵。

## 修改文件

- Abstractions：`ExcelEntityListExtensions.cs`、`ExcelEntityLayout.cs`。
- Provider：`NpoiEntityLayoutExecutors.cs`、`ClosedXmlEntityLayoutExecutor.cs`。
- 测试与消费者：NPOI、ClosedXML 职责级测试、`EntityLayoutProviderContractSuite.cs`、`EntityLayoutReusableContractTest.cs`、`PublicApiContractTest.cs`、第三方消费者 `Program.cs`。
- 治理与文档：双 TFM API 基线、成员审批、生产符号追溯、阶段索引和上述两份能力文档。

## API/数据/配置变化

- net6.0、net8.0 相对上一批准基线各新增四个公开成员：`FormulaSumDetailRowsAbove`、`ExcelEntityFooterDetailSumValue`、`SourceColumn`、`ToFormulaA1`。无公共成员删除或签名变化。
- 不改变文件格式、持久化数据、默认 Provider 选择或版本号。

## 测试结果

- NPOI 新能力定向测试：net6.0/net8.0 各 9/9；ClosedXML 各 5/5；共享合同定向测试各 3/3。
- Release 常规矩阵（`Category!=Large`）：除首次运行时尚未更新的 API 治理门禁外，各项目测试均通过；更新类型分类和基线后，集成测试 net6.0/net8.0 各 32/32 通过。合计 3314 项已通过。
- 本地重新打包五个生产 NuGet 包及 Entity Provider 合同包，第三方消费者在 net6.0、net8.0 均输出 `third-party-public-only-provider-ok`。

## Build/Typecheck/Lint/Format

- `Bing.Offices.sln` Release 构建：0 错误。现有 net6.0 依赖支持、NuGet 源与依赖解析警告保留。
- `git diff --check`：无空白错误。修改文件均通过严格 UTF-8 解码；C# 为 BOM+LF，Markdown 为无 BOM+LF，API 基线保留原有 CRLF。
- API 快照对比：双 TFM 各新增计划内四个成员，无删除；成员级审批与类型分类已同步。

## 计划偏差

首次常规矩阵中的六项 API 治理失败属于尚未写入新成员分类和快照的执行顺序问题，补齐后对应双 TFM 集成测试全部通过。其余按计划执行。

## 基线问题

工作区原有较大未提交改动和若干构建警告；未清理、覆盖或提交这些改动。

## 已知问题

无本阶段未解决的功能问题。公式由 Excel 计算，本库不缓存计算结果。

## 风险与回归关注点

跨大量分页或分组产生的 SUM 参数会分段为多个 SUM，超过 Excel 公式长度上限时明确失败；该失败不应覆盖已存在的目标文件。

## Reviewer 注意事项

重点核对明细行段一基坐标、分页与分组小计后的最终 Footer 定位，以及 Provider SPI 的四项新增成员。工作区整体差异包含此前阶段的改动，不应归因于本阶段。

## Git 状态

未自动 git commit、push、创建 PR、发布或修改版本号。
