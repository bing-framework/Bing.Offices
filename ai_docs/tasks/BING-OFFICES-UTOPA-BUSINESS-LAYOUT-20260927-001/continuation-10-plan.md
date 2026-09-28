# Continuation 10：跨页签字区

用户授权继续后续阶段。本轮落实推荐顺序中的跨页签字区；复用已公开的 PageBreak、PageSubtotal 和 Footer，不增加同义 API。

## 实施范围

1. 中间页的小计和签字区使用多行 PageSubtotal，末页总计及签字区使用 Footer；支持 GapRows、相对合并、模板样式和空明细。
2. 修复分页额外行占用后的末页明细边界预检；拒绝末页后多余的小计 marker，保持 Plan 结构化错误及目标流完整性。
3. NPOI XLS/XLSX 与 ClosedXML XLSX 直接测试完整内容、合并、分页符、模板保留、同步/异步导入导出、有无表头、空/整页/多页明细；补齐声明和物理边界回归。
4. 完善可执行文档示例和生产符号追溯，记录真实验证结果及既存编码偏差。

## 影响及验证

- ChangedProjects：NPOI、ClosedXML 及其职责测试、Docs。
- ChangedPublicContracts：无；API 成员快照须保持不变，不覆盖基线。
- ChangedRuntimePaths：Entity 分页预检及导入 marker 校验。
- ChangedProviders/TFMs：NPOI、ClosedXML；net6.0/net8.0。
- ChangedBuildPackaging/BenchmarkHarness：无。
- RiskLevel：MEDIUM；先 L1 新增场景，再双 TFM Provider L2、Provider Contract L3 和 Docs；阶段收口构建 solution。
- 不重跑无关大规模性能矩阵。没有新增 API，不重新批准或覆盖 API baseline。

## 后续

尾部命名锚点需独立确定重复页锚点命名、模板名称作用域及动态地址更新契约；本轮不将其与签字区表达能力混为一项。后续依次推进尾部命名锚点、可复用 Provider 契约测试模板、静态分析器。
