# 报表级工作簿导出策略

Task ID: `BING-OFFICES-WORKBOOK-EXPORT-STRATEGY-20260928-001`

## 目标

调用方按报表显式选择完整工作簿导出或前向流式导出。策略入口通过现有公开接口分发，按 Provider 声明预检格式和能力；请求不兼容时在目标流/文件被写入前失败，不自动切换 Provider 或模式。

## 边界

- 完整工作簿模式调用 `IExcelExporter`，不保证 Provider 内部始终使用 DOM；NPOI 对大型纯列表仍可使用 SXSSF，MiniExcel 仍使用自身前向写入实现。
- 前向流式模式调用 `IExcelStreamingExporter`，禁止模板。Provider 继续负责其详细布局限制。
- 保留原接口、构造函数、文件提交、取消、异常观察和流所有权合同。
- 不改变 Entity Layout 与导入 API，不修改 Provider 的内部导出算法。

## 实施

1. 在 Core 增加 `ExcelWorkbookExportMode` 与 `ExcelWorkbookExportStrategy`，提供流/文件的同步、异步方法。构造函数接收两个可选 Provider 实例；仅所选模式的 Provider 必须存在。
2. 预检模式、请求、目标、格式、模板和已声明的 Provider 能力。未知模式、缺失 Provider 或能力不匹配时失败，且不调用另一模式。
3. 增加职责级测试：NPOI/ClosedXML 完整工作簿，SpreadCheetah 流式，格式与模板拒绝、缺失 Provider、取消、输出流所有权及文件提交路径。
4. 更新 Provider 文档、API 分类/双 TFM 快照、生产符号到测试方法追溯和 NuGet 消费示例。
5. 先执行职责测试与受影响项目，再执行 API/消费者/文档和阶段检查；保留原有未提交改动及编码行尾。

## 验收

模式选择明确，结果工作簿可读；不兼容请求在任何输出前失败；同步与异步只进入所选 Provider；旧 API 无删除，双 TFM 公共 API 差异仅为本阶段新增成员。
